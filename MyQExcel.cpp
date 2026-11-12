#include "MyQExcel.h"

#include <QCryptographicHash>
#include <QFile>
#include <QDir>
#include <QMessageBox>

#include "MyQShortings.h"
#include "MyQByteArray.h"

struct axuptr
{
	QAxObjectUptr uptr;

	axuptr() noexcept = default;
	axuptr(std::nullptr_t) noexcept {}
	axuptr(QAxObject* ptr) noexcept : uptr(ptr) {}

	// Разрешаем только перемещение, запрещаем случайное копирование
	axuptr(axuptr&&) noexcept = default;
	axuptr& operator=(axuptr&&) noexcept = default;
	axuptr(const axuptr&) = delete;
	axuptr& operator=(const axuptr&) = delete;

	// Безопасное приведение к bool (защита от неявного каста в int)
	explicit operator bool() const noexcept { return static_cast<bool>(uptr); }

	QAxObject* operator->() const noexcept { return uptr.get(); }
	QAxObject& operator*() const noexcept { return *uptr; }

	// Удобный доступ к сырому указателю (например, для передачи в dynamicCall)
	QAxObject* get() const noexcept { return uptr.get(); }
	operator QAxObject*() const noexcept { return uptr.get(); }
};

QString MyQExcelRows::ToString()
{
	QString ret;

	for(uint r=0; r<data.size(); r++)
	{
		for(int c=0; c<data[r].size(); c++)
		{
			ret += data[r][c] + ISheet::EndCell;
		}
		ret += ISheet::EndRow;
	}

	return ret;
}


// Преобразует координаты с нуля (x - колонка, y - строка) в адрес ячейки Excel (например, 0, 0 -> "A1")
QString GetExcelAddr(int x, int y) {
	QString colStr;
	int col = x + 1; // Переводим в 1-базную систему Excel
	while (col > 0) {
		colStr.prepend(QChar('A' + (col - 1) % 26));
		col = (col - 1) / 26;
	}
	return colStr + QString::number(y + 1);
}

ExcelSheet::ExcelSheet(QAxObject &aSheet, IWorkbook &parent):
	parent{parent},
	sheet{aSheet}
{

}

QSize ExcelSheet::Size() const
{
	QSize size;

	axuptr usedRange = sheet.querySubObject("UsedRange");
	if(not usedRange)
	{
		parent.Error("null usedRange");
		return size;
	}

	int firstRow = usedRange->property("Row").toInt();
	int firstCol = usedRange->property("Column").toInt();

	axuptr rows = usedRange->querySubObject("Rows");
	axuptr cols = usedRange->querySubObject("Columns");

	if (rows && cols) {
		size.setHeight(firstRow + rows->property("Count").toInt() - 1);
		size.setWidth(firstCol + cols->property("Count").toInt() - 1);
	}
	else parent.Error("null rows || cols");

	return size;
}

QString ExcelSheet::Cell(int col, int row) const
{
	axuptr cell = sheet.querySubObject("Cells(int,int)", row+1, col+1);

	if(!cell) { parent.Error("null cell"); return ""; }

	return cell->property("Value").toString();
}

void ExcelSheet::SetCell(int col, int row, const QString &value)
{
	axuptr cell = sheet.querySubObject("Cells(int,int)", row+1, col+1);

	if(!cell) { parent.Error("null cell"); return; }

	cell->setProperty("Value",value);
}

ExcelSheet::Rows ExcelSheet::Range(const QRect &rect) const
{
	QString addr = GetExcelAddr(rect.x(), rect.y()) + ":" + GetExcelAddr(rect.x()+rect.width(), rect.y()+rect.height());

	axuptr range = sheet.querySubObject("Range(const QString&)", addr);
	if (!range) { parent.Error("null range"); return {}; }

	// 2. Забираем все данные одним вызовом
	QVariant var = range->property("Value2");

	Rows rows1;

	// 3. Разбираем результат
	// Если диапазон состоял из 1 ячейки (x1 == x2 && y1 == y2), вернется одиночный QVariant, а не список
	if (var.userType() != QMetaType::QVariantList) {
		auto &row = rows1.data.emplace_back();
		row.push_back(var.toString());
	}
	// Если диапазон состоял из более 1 ячейки
	else
	{
		QVariantList varsList = var.toList();
		for (int r = 0; r < varsList.size(); ++r) {
			QVariantList cols = varsList[r].toList();
			auto &row = rows1.data.emplace_back();
			for (int c = 0; c < cols.size(); ++c) {
				row.push_back(cols[c].toString());
			}
		}
	}

	return rows1;
}

ExcelSheet::Rows ExcelSheet::RangeAllSheet() const
{
	auto size = Size();
	return Range(QRect(0, 0, size.width(), size.height()));
}

void ExcelSheet::ClearContentsSheet()
{
	axuptr cells = sheet.querySubObject("Cells");
	if (!cells) { parent.Error("cells is null"); return; }

	cells->dynamicCall("ClearContents()");
}

ExcelWorkbook::ExcelWorkbook(ErrorWorker errorWorker):
	IWorkbook(std::move(errorWorker))
{

}

ExcelWorkbook::ExcelWorkbook(QString file, ErrorWorker errorWorker):
	IWorkbook(std::move(errorWorker))
{
	auto error = Open(file);
	if(!error.isEmpty()) Error(error);
}

ExcelWorkbook::~ExcelWorkbook() {
	if(IsOpened()) CloseNoSave();
}

QString ExcelWorkbook::Open(QString file)
{
	/// работу с QAxObject не требуется оборачивать в try-catch поскольку они не выбрасывают исключения,
	/// а при ошибках отправляют сигналы

	/// Вызовы querySubObject возвращают указатель на объект в куче!!! Нужно удалять!
	/// Удобно использовать RAII axuptr!

	if(!errorWorker)
	{
		Error("ExcelWorkbook::Open, but errorWorker is null");
	}

	if(excel_uptr) { ClearFields(); return "ExcelWorkbook::Open is aleady opened"; }
	if(!sheets.empty()) { ClearFields(); return "ExcelWorkbook::Open sheets not empty"; }

	if(not QFileInfo(file).isFile()) return file + " is not file";

	excel_uptr = std::make_unique<QAxObject>("Excel.Application");

	auto workbooks = excel_uptr->querySubObject("Workbooks"); // будет удалён за счёт парента
	if(!workbooks) { ClearFields(); return "ExcelWorkbook::Open can't get workbooks"; }

	workbook = workbooks->querySubObject("Open(const QString&)", file);
	if(!workbook) { ClearFields(); return "ExcelWorkbook::Open can't get workbook"; }

	auto __sheets = workbook->querySubObject("Worksheets");  // будет удалён за счёт парента
	if(!__sheets) { ClearFields(); return "ExcelWorkbook::Open null sheets"; }

	int shC = __sheets->property("Count").toInt();
	for(int i=1; i<=shC; i++)
	{
		auto axSheet = __sheets->querySubObject("Item(int)", i); // будет удалён за счёт парента
		if(!axSheet) { ClearFields(); return "ExcelWorkbook::Open can't get axSheet"; }

		auto sheet = std::make_unique<ExcelSheet>(*axSheet, *this);

		this->sheets.push_back(std::move(sheet));
	}

	return "";
}

QString ExcelWorkbook::Save()
{
	if(!workbook) return "ExcelWorkbook::Save called for null workbook";

	workbook->dynamicCall("Save()"); // всегда возвращает пустой вариант, проверять нечего

	return "";
}

QString ExcelWorkbook::CloseNoSave()
{
	if(!excel_uptr) return "ExcelWorkbook::CloseNoSave called for null excel_uptr";
	if(!workbook) return "ExcelWorkbook::CloseNoSave called for null workbook";

	sheets.clear();

	workbook->dynamicCall("Close(bool)", false);
	excel_uptr->dynamicCall("Quit()");
	ClearFields();

	return "";
}

ISheet * ExcelWorkbook::Sheet(int index) const {
	if(index >= 0 and index < (int)sheets.size())
		return sheets[index].get();

	return nullptr;
}

QString ExcelWorkbook::ToStr() const
{
	QString ret;
	int shC = SheetCount();
	for(int shi = 0; shi<shC; shi++)
	{
		ISheet *sheet = Sheet(shi);
		ret += sheet->Name() +  ISheet::EndSheetName;
		ret += sheet->RangeAllSheet().ToString() + ISheet::EndSheet;
	}

	return ret;
}



void QStringsSheet::Init(const QString &sheet)
{
	rows.clear();
	name = sheet.left(sheet.indexOf(EndSheetName));
	QStringList rows_ = sheet.split(EndRow);
	if(rows_.size())
	{
		rows_[0].remove(0,rows_[0].indexOf(EndSheetName) + EndSheetName.length());
		for(auto &r:rows_)
		{
			rows.push_back(r.split(EndCell));
		}
	}
}

QString QStringsSheet::Cell(int col, int row) const
{
	throw std::logic_error("QStringsSheet::Cell");
	//if(row >= (int)rows.size() || col >= rows[row].size())
		//parent.Error("QString Cell(col=" + QSn(col) + ", row=" + QSn(row) + ")  out of boud");
	return rows[row][col];
}

ExcelQStringWorkbook::ExcelQStringWorkbook(IWorkbook::ErrorWorker errorWorker):
	IWorkbook(std::move(errorWorker))
{

}

ISheet* ExcelQStringWorkbook::Sheet(int number) const
{
	if(number > (int)sheetsPt.size())
		Error("IExcelSheet* Sheet(number=" + QSn(number) + ") out of boud");
	return sheetsPt[number-1];
}

QString ExcelQStringWorkbook::ToStr() const
{
	QString ret;
	for(auto &sh:sheets)
	{
		ret += sh.Name() +  ISheet::EndSheetName;
		ret += sh.RangeAllSheet().ToString() + ISheet::EndSheet;
	}
	return ret;
}

void ExcelQStringWorkbook::FromStr(const QString &/*workBookAsStr*/)
{
	IWorkbook::Error("ExcelQStringWorkbook::FromStr отключено, нужно реализовать вирт функцию");
//	QStringList sheets_ { workBookAsStr.split(IExcelSheet::EndSheet()) };
//	for(int i=0; i<sheets_.size(); i++)
//		sheets.push_back(sheets_[i]);

//	for(uint i=0; i<sheets.size(); i++)
//		sheetsPt.push_back(&sheets[i]);
}

QString MyQExcelHasher::ErrCodeToStr(int errCode)
{
	if(errCode == noError) return "No error";
	if(errCode == cantCrDir) return "Ошибка создания директории для файлов хеширования";
	if(errCode == cantOpenExcelFile) return "cant Open Excel File";
	if(errCode == cantOpenFileHashes) return "Can't open file /_hashes.txt";
	if(errCode == cantOpenFileHash) return "Can't open file hash";
	if(errCode == cantOpenFileNew) return "Can't open new file";
	if(errCode == cantOpenNewFileHash) return "Can't open new file hash";
	return "Wrong errCode "+QSn(errCode);
}

QString MyQExcelHasher::GetFileName(int fileIndex)
{
	QString indStr = QSn(fileIndex);
	while(indStr.length()<6) indStr = "0" + indStr;
	return "file_" + indStr + ".txt";
}

int MyQExcelHasher::LoadFromHashedFile(QString fileName)
{
	Log("Load from hashed file " + QFileInfo(fileName).fileName());
	QFile fileData(fileName);

	if(fileData.open(QFile::ReadOnly))
	{
		QString wbStr = fileData.readAll();
		fileData.close();

		ExcelQStringWorkbook *swb = new ExcelQStringWorkbook([](const QString &text){ QMbError(text); });
		wbPt = swb;
		swb->FromStr(wbStr);
	}
	else return cantOpenFileHash;

	return noError;
}

int MyQExcelHasher::InitFile(QString directory, QString fileName)
{
//	labelForMsgs = &labelForMsgs_;
	QDir dirHashedFiles(directory + "/hashed_files");
	if(!dirHashedFiles.exists()) dirHashedFiles.mkdir(dirHashedFiles.path());
	if(!dirHashedFiles.exists()) return cantCrDir;

	QCryptographicHash curHash(QCryptographicHash::Md5);
	QFile file(fileName);
	file.open(QFile::ReadOnly);
	curHash.addData(&file);
	file.close();

	QStringList hashesFileContent;
	QFile file_hashes(dirHashedFiles.path() + "/_hashes.txt");
	if(file_hashes.exists())
	{
		if(file_hashes.open(QFile::ReadWrite))
		{
			QString hashes = file_hashes.readAll();
			file_hashes.close();

			hashesFileContent = hashes.split("[endHashLine]\r\n", QString::SkipEmptyParts);
		}
		else return cantOpenFileHashes;
	}
	else { /*if file_hashes doesn't exists it's not error, it will be created later*/ }

	QString findFile;
	int fileIndex = 0;
	for(auto &hashFileLine:hashesFileContent)
	{
		QStringList hash1LineElements {hashFileLine.split("[endHashElem]")};
		QString hash = hash1LineElements[0];
		fileIndex = hash1LineElements[1].toInt();
		if(hash == MyQBA::QByteArrToStr(curHash.result())) findFile = GetFileName(fileIndex);
	}

	if(findFile != "")
	{
		return LoadFromHashedFile(dirHashedFiles.path() + "/" + findFile);
	}
	else
	{
		fileIndex++;
		QString newFile = GetFileName(fileIndex);
		Log("Hash not found, write new hash " + newFile);

		ExcelWorkbook ewb([](const QString &text){ QMbError(text); });
		if(ewb.Open(fileName).isEmpty())
		{
			if(file_hashes.open(QFile::Append))
			{
				QString curHashStr = MyQBA::QByteArrToStr(curHash.result());
				QTextStream stream(&file_hashes);
				stream << curHashStr << "[endHashElem]";
				stream << QSn(fileIndex) << "[endHashElem]";
				stream << fileName << "[endHashElem][endHashLine]\r\n";
				file_hashes.close();

				QFile fileData(dirHashedFiles.path() + "/" + newFile);
				if(fileData.open(QFile::WriteOnly))
				{
					QTextStream stream(&fileData);
					stream << ewb.ToStr();
					fileData.close();

					ewb.CloseNoSave();

					return LoadFromHashedFile(dirHashedFiles.path() + "/" + newFile);
				}
				else return cantOpenNewFileHash;
			}
			else return cantOpenFileHashes;
		}
		else return cantOpenExcelFile;
	}

	return noError;
}
