//--------------------------------------------------------------------------------------------------------------------------
#ifndef MYQDIFFERENT_H
#define MYQDIFFERENT_H
//--------------------------------------------------------------------------------------------------------------------------
#include <set>

#include <QCoreApplication>
#include <QDebug>
#include <QFile>
#include <QFileInfo>
#include <QWidget>
#include <QLineEdit>
#include <QDir>
#include <QDateTime>

#include "MyQShortings.h"
#include "MyQFileDir.h"
//--------------------------------------------------------------------------------------------------------------------------

struct Time_dhms;

struct MyQDifferent
{
	inline static QString PathToExe() { return ExePath(); }
	inline static QString ExePath();
	inline static QString ExeNameNoPath();
	inline static QString ExeNameWithPath();

	inline static QString BytesToString(uint64_t bytesCount);

	inline static void GetPathName(const QString &file, QString *path, QString *name);
	inline static std::pair<QString, QString> GetPathName(const QString &file);

	inline static QStringList ArgsToStrList(int argc, char *argv[]);
	
	inline static void SetSelectedText(QLineEdit *lineEdit, const QString &newText);

	inline static Time_dhms CalcSeconds(uint64_t seconds);
};

//------------------------------------------------------------------------------------------------------------------------------------------

struct Time_dhms
{
	uint64_t d = 0;
	uint h = 0;
	uint m = 0;
	uint s = 0;

	inline Time_dhms() = default;
	inline Time_dhms(uint64_t seconds);

	inline QString ToString() const;
};

//------------------------------------------------------------------------------------------------------------------------------------------

QString MyQDifferent::ExePath()
{
	// альтернативой applicationDirPath() является QDir::currentPath(), но если программа запущена из Qt игнорирует /debug
	return QCoreApplication::applicationDirPath();
}

QString MyQDifferent::ExeNameNoPath()
{
	return QFileInfo(QCoreApplication::applicationFilePath()).fileName();
}

QString MyQDifferent::ExeNameWithPath()
{
	return QCoreApplication::applicationFilePath();
}

QString MyQDifferent::BytesToString(uint64_t bytesCount)
{
	if(bytesCount > (uint64_t)1024 * 1024 * 1024 * 1024)
		return QSn(float(bytesCount) / float((uint64_t)1024 * 1024 * 1024 * 1024), 'f', 1) + " Tb";
	if(bytesCount > 1024 * 1024 * 1024) return QSn(float(bytesCount) / float(1024 * 1024 * 1024), 'f', 1) + " Gb";
	if(bytesCount > 1024 * 1024       ) return QSn(float(bytesCount) / float(1024 * 1024       ), 'f', 1) + " Mb";
	if(bytesCount > 1024              ) return QSn(float(bytesCount) / float(1024              ), 'f', 1) + " Kb";
	return QSn(bytesCount) + " b";
}

void MyQDifferent::GetPathName(const QString &file, QString *path, QString *name)
{
	QFileInfo fi(file);
	*path = fi.path();
	*name = fi.fileName();
}

std::pair<QString, QString> MyQDifferent::GetPathName(const QString &file)
{
	std::pair<QString, QString> res;
	GetPathName(file, &res.first, &res.second);
	return res;
}

QStringList MyQDifferent::ArgsToStrList(int argc, char *argv[])
{
	QStringList args;
	for(int i=0; i<argc; i++)
	{
		args += QString::fromLocal8Bit(argv[i]);
	}
	return args;
}

void MyQDifferent::SetSelectedText(QLineEdit *lineEdit, const QString &newText)
{
	if (!lineEdit->hasSelectedText()) return;
	QString currentText = lineEdit->text();
	int start = lineEdit->selectionStart();
	currentText.replace(start, lineEdit->selectionLength(), newText);
	lineEdit->setText(currentText);
	lineEdit->setSelection(start, newText.length());
}

Time_dhms MyQDifferent::CalcSeconds(uint64_t seconds)
{
	return Time_dhms(seconds);
}

//--------------------------------------------------------------------------------------------------------------------------

Time_dhms::Time_dhms(uint64_t seconds):
	d { uint(seconds / 86400) },
	h { uint((seconds % 86400) / 3600) },
	m { uint((seconds % 3600) / 60) },
	s { uint(seconds % 60) }
{}

QString Time_dhms::ToString() const
{
	//		QString res;
	//		if(d>0) res += QSn(d)+"d";
	//		if(h>0 or !res.isEmpty()) res += QSn(h)+"h";
	//		if(m>0 or !res.isEmpty()) res += QSn(m)+"m";
	//		if(s>0 or !res.isEmpty()) res += QSn(s)+"s";
	//		return res;

	if(d == 0 and h == 0 and m == 0 and s == 0) return "zero time";
	QString res;

	if(d>0) res += d == 1 ? QSn(d)+" day"	// d == 1
						  : QSn(d)+" days";	// d > 1

	if(h == 0 and m == 0 and s == 0) res += " exactly";
	else
	{
		if(!res.isEmpty()) res.append(" + ");

		res.append(QTime(h,m,s).toString(TimeFormat));
	}
	return res;
}

//--------------------------------------------------------------------------------------------------------------------------
#endif
//--------------------------------------------------------------------------------------------------------------------------



