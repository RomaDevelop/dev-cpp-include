#ifndef MyQTimer_H
#define MyQTimer_H

#include <set>
#include <queue>
#include <memory>
#include <functional>

#include <QDebug>
#include <QTimer>
#include <QDateTime>

//--------------------------------------------------------------------------------------------------------------------------

///\brief MyQTimerPool is intended for use in one thread.
struct MyQTimerPool
{
	MyQTimerPool() { }
	~MyQTimerPool() { ClearTimers(); }

	inline static void SingleShotSt(QObject *parent, int msec, std::function<void()> task);

	using Tasks = std::vector<std::pair<int, std::function<void()>>>;
	inline static void SingleShotsSt(QObject *parent, Tasks tasks);

	inline void SingleShot(QObject *parent, int msec, std::function<void()> function);
	inline void SingleShots(QObject *parent, Tasks tasks);
	inline void ClearTimers();

private:
	std::set<QTimer*> allTimers;
	std::set<QTimer*> freeTimers;

	static MyQTimerPool& staticPool() { static MyQTimerPool pool; return pool; }
};

struct MyQTimerPoolSt
{
	inline void SingleShot(QObject *parent, int msec, std::function<void()> function) { MyQTimerPool::SingleShotSt(parent, msec, std::move(function)); }
	using Tasks = MyQTimerPool::Tasks;
	inline void SingleShots(QObject *parent, Tasks tasks){ MyQTimerPool::SingleShotsSt(parent, std::move(tasks)); }
};

//--------------------------------------------------------------------------------------------------------------------------

class MyQTimer : public QTimer {
    Q_OBJECT
public:

	///\brief Асинхронный цикл выполняемый в потоке GUI
	/// эквивалент for (int i = i_start; i < i_less_than; i++), но итерации выполняются через интервалы interval_ms
	/// агрумент iteration должен быть не пуст
	/// to_do: Сейчас таймеру просто ставится интервал, нужно сделать версию с паузами между итерациями,.
	///		   Текущая версия отлично работает с нетяжётыми операциями, но если iteration будет выполнятся дольше interval_ms.
	inline static void For(int i_start, int i_less_than, int interval_ms, QObject &parent,
	                       std::function<void(int &i)> iteration,
	                       std::function<void()> finishAction = {});

	///\brief Отложенный вызов функции по времени или условию
	/// Создает таймер с интервалом step_ms,
	/// вызывает call, если checkFn вернуло true или прошло больше deferral_ms
	inline static void DeferredCall(int deferral_ms, int step_ms, QObject &parent,
							 std::function<bool()> checkFn,
							 std::function<void()> call);

	int timerTicks = 0;
	int counter = 0;
	int tag = 0;
	int finish = false;
	QString stringValue;
	std::map<QString,int> intValues;
	explicit MyQTimer(QObject *parent = nullptr) : QTimer(parent) { }
	void Finish(bool deleteTimer)
	{
		stop();
		if(deleteTimer) deleteLater();
	}
};

void MyQTimer::For(int i_start, int i_less_than, int interval_ms, QObject &parent,
				   std::function<void (int &)> iteration, std::function<void ()> finishAction)
{
	if (i_start >= i_less_than) // необходимо, иначе если передать i_start >= i_less_than, iteration выполнится 1 раз
	{
		if (finishAction) finishAction();
		return;
	}

	/// Изначально i передавал в лямбду в shared_ptr, но оказывается это не нужно, все работает так.
	/// (в mutable лямбде захваченная переменная i = i_start будет сохранять свое значение между вызовами)
	QTimer *timer = new QTimer(&parent);
	timer->setSingleShot(true); // Включен режим SingleShot, поэтому вызывать stop по окончанию не потребуется

	/// Используется mutable-лямбда, т.е. счётчик i будет сохранять свое значение между срабатываниями таймена
	connect(timer, &QTimer::timeout, timer,
	        [
	        i = i_start,
	        i_less_than,
	        interval_ms,
	        iteration = std::move(iteration),
	        finishAction = std::move(finishAction),
	        timer
	        ]() mutable
	{
		iteration(i);
		i++;

		if(i < i_less_than)
		{
			timer->start(interval_ms);
		}
		else
		{
			if(finishAction) finishAction();
			// Вызов timer->stop() - не треуется поскольку таймер в режиме SingleShot
			timer->deleteLater();
		}
	});

	timer->start(interval_ms);
}

void MyQTimer::DeferredCall(int deferral_ms, int step_ms, QObject &parent,
					std::function<bool ()> checkFn, std::function<void ()> call)
{
	QTimer *timer = new QTimer(&parent);
	QDateTime start = QDateTime::currentDateTime();
	connect(timer, &QTimer::timeout, &parent, [checkFn, call, timer, start, deferral_ms](){
		if(checkFn() or start.msecsTo(QDateTime::currentDateTime()) > deferral_ms)
		{
			call();
			timer->stop();
			timer->deleteLater();
		}
	});
	timer->start(step_ms);
}

void MyQTimerPool::SingleShotSt(QObject *parent, int msec, std::function<void ()> task)
{
	staticPool().SingleShot(parent, msec, std::move(task));
}

void MyQTimerPool::SingleShotsSt(QObject *parent, MyQTimerPool::Tasks tasks)
{
	staticPool().SingleShots(parent, std::move(tasks));
}

void MyQTimerPool::SingleShot(QObject *parent, int msec, std::function<void ()> function)
{
	QTimer *timer;
	if(freeTimers.empty())
	{
		timer = new QTimer(parent);
		timer->setSingleShot(true);
		allTimers.insert(timer);
		QObject::connect(timer, &QObject::destroyed, [this, timer](){ freeTimers.erase(timer); allTimers.erase(timer); });
	}
	else
	{
		timer = *freeTimers.begin();
		freeTimers.erase(freeTimers.begin());
		timer->setParent(parent);
	}
	QObject::connect(timer, &QTimer::timeout, [this, timer, function = std::move(function)](){
		function();
		timer->setParent(nullptr);
		QObject::disconnect(timer, &QTimer::timeout, nullptr, nullptr);
		freeTimers.insert(timer);
	});
	timer->start(msec);
}

void MyQTimerPool::SingleShots(QObject *parent, MyQTimerPool::Tasks tasks)
{
	for(auto &[msec, function]:tasks)
	{
		SingleShot(parent, msec, std::move(function));
	}
}

void MyQTimerPool::ClearTimers()
{
	std::vector<QTimer*> timers { allTimers.begin(), allTimers.end() };
	// should work with copy of list, because original list changes when timer destroys

	for(auto &timer:timers)
	{
		timer->deleteLater();
	}
}

//--------------------------------------------------------------------------------------------------------------------------
#endif
//--------------------------------------------------------------------------------------------------------------------------
