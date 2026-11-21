#ifndef MyQTimer_H
#define MyQTimer_H

#include <set>
#include <queue>
#include <memory>
#include <functional>

#include <QDebug>
#include <QEventLoop>
#include <QTimer>
#include <QDateTime>

//--------------------------------------------------------------------------------------------------------------------------

///\brief Пул SingleShot таймеров для оптимизации выполнения отложенных задач в основном потоке.
///\warning Класс предназначен для работы в одном потоке.
struct MyQTimerPool
{
	MyQTimerPool() { }
	~MyQTimerPool() { ClearTimers(); }

	///\brief Задача task должна быть выполнена через interval_ms
	struct Task { uint interval_ms = 0; std::function<void()> task; };
	using Tasks = std::vector<Task>;

	///\brief Выделяет таймер из пула для выполнения задачи task через interval_ms
	inline void SingleShot(QObject *parent, int msec, std::function<void()> function);
	///\brief Выделяет циклом таймеры из пула для выполнения задач tasks
	inline void SingleShots(QObject *parent, Tasks tasks);
	///\brief Уничтожает все таймеры в пуле вызовом deleteLater для каждого
	inline void ClearTimers();

	///\brief Выделяет таймер из статического пула для выполнения задачи task через interval_ms
	inline static void SingleShotSt(QObject *parent, int interval_ms, std::function<void()> task);
	///\brief Выделяет циклом таймеры из статического пула для выполнения задач tasks
	inline static void SingleShotsSt(QObject *parent, Tasks tasks);
	///\brief Уничтожает все таймеры в статическом пуле вызовом deleteLater для каждого
	inline static void ClearTimersSt();

private:
	std::set<QTimer*> allTimers;
	std::set<QTimer*> freeTimers;

	static MyQTimerPool& staticPool() { static MyQTimerPool pool; return pool; }
};

///\brief Альтернативный способ работы со статическим пулом таймеров
struct MyQTimerPoolSt
{
	using Tasks = MyQTimerPool::Tasks;

	inline void SingleShot(QObject *parent, int msec, std::function<void()> function);
	inline void SingleShots(QObject *parent, Tasks tasks);
	inline void ClearTimers();
};

//--------------------------------------------------------------------------------------------------------------------------

class MyQTimer : public QTimer {
    Q_OBJECT
public:

	///\brief Асинхронный цикл выполняемый в потоке GUI
	/// Эквивалент for (int i = i_start; i < i_less_than; i++), но итерации выполняются через интервалы interval_ms.
	/// Агрумент iteration должен быть не пуст.
	/// Если цикл будет прерван из-за уничтожения parent, finishAction не будет выполнено.
	inline static void For(int i_start, int i_less_than, int interval_ms, QObject &parent,
	                       std::function<void(int &i)> iteration,
	                       std::function<void()> finishAction = {});

	///\brief Аналогично функции MyQTimer::For, но блокирует выполнение кода ниже до завершения работы цикла.
	/// Использует для блокировки QEventLoop и вызывает MyQTimer::For.
	inline static void ForBlocking(int i_start, int i_less_than, int interval_ms, QObject &parent,
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

void MyQTimer::ForBlocking(int i_start, int i_less_than, int interval_ms, QObject & parent,
                    std::function<void (int &)> iteration, std::function<void ()> finishAction)
{
	QEventLoop loop;

	// Страховка: если parent умрет раньше окончания цикла, loop будет завершен
	QObject::connect(&parent, &QObject::destroyed, &loop, &QEventLoop::quit);

	MyQTimer::For(i_start, i_less_than, interval_ms, parent, std::move(iteration),
	              [finishAction = std::move(finishAction), &loop]()
	{
		if(finishAction) finishAction();
		loop.quit();
	});

	loop.exec();
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
	for(auto &task:tasks)
	{
		SingleShot(parent, task.interval_ms, std::move(task.task));
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

void MyQTimerPool::SingleShotSt(QObject * parent, int interval_ms, std::function<void ()> task)
{
	staticPool().SingleShot(parent, interval_ms, std::move(task));
}

void MyQTimerPool::SingleShotsSt(QObject * parent, MyQTimerPool::Tasks tasks)
{
	staticPool().SingleShots(parent, std::move(tasks));
}

void MyQTimerPool::ClearTimersSt()
{
	staticPool().ClearTimers();
}

void MyQTimerPoolSt::SingleShot(QObject * parent, int msec, std::function<void ()> function)
{
	MyQTimerPool::SingleShotSt(parent, msec, std::move(function));
}

void MyQTimerPoolSt::SingleShots(QObject * parent, MyQTimerPoolSt::Tasks tasks)
{
	MyQTimerPool::SingleShotsSt(parent, std::move(tasks));
}

void MyQTimerPoolSt::ClearTimers()
{
	MyQTimerPool::ClearTimersSt();
}

//--------------------------------------------------------------------------------------------------------------------------
#endif
//--------------------------------------------------------------------------------------------------------------------------
