#pragma once

using TimerCallback = std::function<void()>;

class Timer
{
	NO_COPY_ASSIGN(Timer);

	friend void CALLBACK TimerProc(void* param, BOOLEAN timerCalled);

public:
	Timer(TimerCallback callback = nullptr);
	~Timer();

	void SetCallback(TimerCallback callback);

	// interval:
	// interval in ms
	//
	// immediately:
	// true to call first event immediately
	//
	// once:
	// true to call timed event only once
	//
	bool Start(unsigned int interval, bool immediately = false, bool once = false);
	void Stop();

	bool IsRunning() const;

private:
	void Timer::OnTimedEvent();
	static void CALLBACK TimerProc(void* param, BOOLEAN timerCalled);

private:
	HANDLE m_hTimer{};
	TimerCallback m_callback{};
};
