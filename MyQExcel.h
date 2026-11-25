#ifndef MYQEXCEL_H
#define MYQEXCEL_H

#include <vector>
#include <memory>

#include <QDebug>
#include <QAxObject>
#include <QTextBrowser>

using QAxObject_uptr = std::unique_ptr<QAxObject>;

namespace MyQAx {

	///\brief RAII-обёртка над умным указателем на QAxObject с поддержкой неявного приведения к сырому указателю.
	///
	/// Структура обеспечивает эксклюзивное владение COM-объектом ActiveX через QAxObjectUptr
	/// (запрещая копирование и разрешая только перемещение) и упрощает работу с API Qt/ActiveX
	/// (например, с QAxBase::dynamicCall) за счёт прозрачного преобразования в QAxObject*.
	///
	/// Использование:
	/// axuptr cell = sheet.querySubObject("Cells(int,int)", row, col); // querySubObject создёт объекты в куче.
	/// Созданный объект захвачен в axuptr и будет автоматически уничтожен при выходе cell из области видимости.
	struct axuptr
	{
		QAxObject_uptr uptr;

		axuptr() noexcept = default;
		axuptr(QAxObject* ptr) noexcept : uptr(ptr) { }
		~axuptr() = default;

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
}

struct MyQExcelRows
{
	std::vector<QStringList> data;
	QString ToString();
	QVariantList ToQVariantList() const;
};

class ISheet
{
public:
	virtual ~ISheet() = default;

	virtual QString Name() const = 0;

	/// col, row нумерация с 0
	virtual QString Cell_(int row, int col) const = 0;
	/// col, row нумерация с 0
	virtual void SetCell_(int row, int col, const QString &value) = 0;

	using Rows = MyQExcelRows;

	/// в rect.x и rect.y нумерация с 0
	virtual Rows Range(const QRect &rect) const = 0;
	virtual Rows RangeAllSheet() const = 0;

	/// colFrom, rowFrom нумерация с 0
	virtual void SetRange(const MyQExcelRows &rows, int rowFrom, int colFrom) = 0;

	virtual std::optional<QPoint> Find(const QString& searchText) const = 0;

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
	void Error(const QString &text) const;
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
	virtual QString Cell_(int row, int col) const override;
	/// col, row нумерация с 0
	virtual void SetCell_(int row, int col, const QString &value) override;

	virtual Rows Range(const QRect &rect) const override;
	virtual Rows RangeAllSheet() const override;

	virtual void SetRange(const MyQExcelRows &rows, int rowFrom, int colFrom) override;

	virtual std::optional<QPoint> Find(const QString& searchText) const override;

	virtual void ClearContentsSheet() override;

	void Error(const QString &text) { parent.Error(text); }
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
	QString Cell_(int row, int col) const override;
	/// col, row нумерация с 0
	void SetCell_(int row, int col, const QString &value) override { rows[row][col] = value; }

	virtual Rows Range(const QRect &/*rect*/) const override { throw std::logic_error("QStringsSheet Range"); }
	virtual Rows RangeAllSheet() const override { throw std::logic_error("QStringsSheet RangeAllSheet"); }
};

///\warning Экземпляры класса привязываются к потоку создания (COM STA).
/// Создание, вызов методов и удаление объекта должны выполняться строго в одном потоке.
class ExcelWorkbook: public IWorkbook
{
	// Объект экселя. Если не пуст - значит книга открыта.
	QAxObject_uptr excel_uptr {};
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

	bool Save();

	QString CloseNoSave();

	ISheet* Sheet(int index) const override;
	ISheet* Sheet(const QString &name) const;
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

	void Log(const QString &str);

	enum errors { noError, unknown, cantCrDir, cantOpenExcelFile, cantOpenFileHashes, cantOpenFileHash,
				  cantOpenFileNew, cantOpenNewFileHash };
	static QString ErrCodeToStr(int errCode);
};


#endif // MYQEXCEL_H
