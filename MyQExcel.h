#ifndef MYQEXCEL_H
#define MYQEXCEL_H

#include <vector>
#include <memory>

#include <QDebug>
#include <QAxObject>
#include <QTextBrowser>

#include "MyQShortings.h"

using QAxObjectUptr = std::unique_ptr<QAxObject>;

struct MyQExcelRows
{
	std::vector<QStringList> data;
	QString ToString();
};

class ISheet
{
public:
	virtual ~ISheet() = default;

	virtual QString Name() const = 0;
	/// col, row нумерация с 0
	virtual QString Cell(int col, int row) const = 0;
	/// col, row нумерация с 0
	virtual void SetCell(int col, int row, const QString &value) = 0;

	using Rows = MyQExcelRows;

	virtual Rows Range(const QRect &rect) const = 0;
	virtual Rows RangeAllSheet() const = 0;

	virtual void ClearContentsSheet() = 0;

	inline static QString EndCell = QStringLiteral("[endCell]");
	inline static QString EndRow = QStringLiteral("[endRow]");
	inline static QString EndSheetName = QStringLiteral("[endSheetName]");
	inline static QString EndSheet = QStringLiteral("[endSheet]");
};

///\warning Экземпляры реализаций типа ExcelWorkbook привязываются к потоку создания (COM STA).
/// Создание, вызов методов и удаление такого объекта должны выполняться строго в одном потоке.
class IWorkbook
{
public:
	using ErrorWorker = std::function<void(const QString &text)>;

	IWorkbook(ErrorWorker errorWorker): errorWorker{std::move(errorWorker)} {}
	virtual ~IWorkbook() = default;

	virtual ISheet* Sheet(int index) const = 0;
	virtual int SheetCount() const = 0;
	virtual QString ToStr() const = 0;


	ErrorWorker errorWorker;
	void Error(const QString &text) const
	{
		if(errorWorker) errorWorker(text);
		else
		{
			qdbg << "IWorkbook::ErrorWorker is null";
			qdbg << "Error:" << text;
		}
	}
};

class ExcelSheet: public ISheet
{
	IWorkbook &parent;
	QAxObject &sheet;

public:
	ExcelSheet(QAxObject &aSheet, IWorkbook &parent);

	virtual ~ExcelSheet() override = default;

	virtual QString Name() const override { return sheet.property("Name").toString(); }

	/// возвращает размер от ячейки 0:0 до max:max
	QSize Size() const;

	/// col, row нумерация с 0
	virtual QString Cell(int col, int row) const override;
	/// col, row нумерация с 0
	virtual void SetCell(int col, int row, const QString &value) override;

	virtual Rows Range(const QRect &rect) const override;
	virtual Rows RangeAllSheet() const override;

	void ClearContentsSheet() override;
};

class QStringsSheet: public ISheet
{
	QString name;
	std::vector<QStringList> rows;

public:
	QStringsSheet(const QString &sheet) { Init(sheet); }
	void Init(const QString &sheet);

	virtual ~QStringsSheet() override { };

	QString Name() const override { return name; }
	/// col, row нумерация с 0
	QString Cell(int col, int row) const override;
	/// col, row нумерация с 0
	void SetCell(int col, int row, const QString &value) override { rows[row][col] = value; }

	virtual Rows Range(const QRect &/*rect*/) const override { throw std::logic_error("QStringsSheet Range"); }
	virtual Rows RangeAllSheet() const override { throw std::logic_error("QStringsSheet RangeAllSheet"); }
};

///\warning Экземпляры класса привязываются к потоку создания (COM STA).
/// Создание, вызов методов и удаление объекта должны выполняться строго в одном потоке.
class ExcelWorkbook: public IWorkbook
{
	// Объект экселя. Если не пуст - значит книга открыта.
	QAxObjectUptr excel_uptr {};
	QAxObject* workbook {}; // будет удалён за счёт parent-а
	std::vector<std::unique_ptr<ExcelSheet>> sheets;

	void Clear()
	{
		sheets.clear();
		workbook = {};
		excel_uptr = {};
	}

public:

	explicit ExcelWorkbook(ErrorWorker errorWorker);
	explicit ExcelWorkbook(QString file, ErrorWorker errorWorker);
	~ExcelWorkbook() override;

	ExcelWorkbook(const ExcelWorkbook &) = delete;
	ExcelWorkbook(ExcelWorkbook &&) = delete;
	ExcelWorkbook& operator = (const ExcelWorkbook &) = delete;
	ExcelWorkbook& operator = (ExcelWorkbook &&) = delete;

	// Возвращает ошибку если была
	QString Open(QString file);
	bool IsOpened() { return bool(excel_uptr); }

	QString Save();

	QString CloseNoSave();

	ISheet* Sheet(int index) const override;
	int SheetCount() const override { return sheets.size(); }
	QString ToStr() const override ;
};

class ExcelQStringWorkbook: public IWorkbook
{
	std::vector<QStringsSheet> sheets;
	std::vector<ISheet*> sheetsPt;

public:

	explicit ExcelQStringWorkbook(ErrorWorker errorWorker);

	ISheet* Sheet(int number) const override;
	int SheetCount() const override { return sheets.size(); }
	QString ToStr() const override;
	void FromStr(const QString &workBookAsStr);
};

class MyQExcelHasher
{
	IWorkbook *wbPt {nullptr};
	QString *labelForMsgs = nullptr;

	static QString GetFileName(int fileIndex);

	int LoadFromHashedFile(QString fileName);

public:
	~MyQExcelHasher() { if(wbPt) delete wbPt; }

	IWorkbook* Get() { return wbPt; }

	int InitFile(QString directory, QString fileName);

	void Log(const QString &str)
	{
		qDebug() << str;
		if(labelForMsgs) *labelForMsgs = str;
	}

	enum errors { noError, unknown, cantCrDir, cantOpenExcelFile, cantOpenFileHashes, cantOpenFileHash,
				  cantOpenFileNew, cantOpenNewFileHash };
	static QString ErrCodeToStr(int errCode);
};


#endif // MYQEXCEL_H
