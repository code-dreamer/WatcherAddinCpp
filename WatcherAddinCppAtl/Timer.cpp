#include "stdafx.h"
#include "Timer.h"

Timer::Timer(TimerCallback callback)
	: m_callback{callback}
{}

Timer::~Timer()
{
	Stop();
}

bool Timer::Start(unsigned int interval, bool immediately, bool once)
{
	if (m_hTimer) {
		return false;
	}

	BOOL success = CreateTimerQueueTimer(&m_hTimer,
										 NULL,
										 TimerProc,
										 this,
										 immediately ? 0 : interval,
										 once ? 0 : interval,
										 WT_EXECUTEINTIMERTHREAD);

	return(success != 0);
}

void Timer::Stop()
{
	DeleteTimerQueueTimer(NULL, m_hTimer, NULL);
	m_hTimer = nullptr;
}

void Timer::SetCallback(TimerCallback callback)
{
	m_callback = callback;
}

void CALLBACK Timer::TimerProc(void* param, BOOLEAN UNUSED(timerCalled))
{
	Timer* timer = static_cast<Timer*>(param);
	timer->OnTimedEvent();
};

void Timer::OnTimedEvent()
{
	if (m_callback)
		m_callback();
}

bool Timer::IsRunning() const
{
	return (m_hTimer != nullptr);
}
