from src.scheduler.clients_scheduler import job_notify_and_cleanup, create_scheduler, default_listener
from apscheduler.events import JobExecutionEvent, EVENT_JOB_EXECUTED, EVENT_JOB_ERROR  # type: ignore
from apscheduler.schedulers.background import BackgroundScheduler  # type: ignore
from unittest.mock import patch, MagicMock
import time


def test_scheduler_adds_jobs_correctly() -> None:
    mocked_job = MagicMock()
    scheduler = BackgroundScheduler()
    scheduler.add_job(func=mocked_job, trigger="interval", seconds=1, args=["test", 123])
    scheduler.start()
    try:
        time.sleep(3)
    finally:
        scheduler.shutdown(wait=False)

    assert mocked_job.called
    mocked_job.assert_called_with("test", 123)

def test_create_scheduler() -> None:
    scheduler = create_scheduler()
    assert isinstance(scheduler, BackgroundScheduler)

    scheduler.start()
    try:
        time.sleep(3)
        jobs = scheduler.get_jobs()
        assert len(jobs) == 1

    finally:
        scheduler.shutdown(wait=False)

def test_job_notify_and_cleanup() -> None:
    with patch('src.scheduler.clients_scheduler.create_client_service') as mock_create_client_service:
        mock_client_service = MagicMock()
        mock_create_client_service.return_value = mock_client_service

        mock_client_service.remove_overdue_clients.return_value = ["client1@example.com"]

        job_notify_and_cleanup()

        mock_client_service.notify_payment_due_in_days.assert_called_once_with(days_ahead=1)
        mock_client_service.remove_overdue_clients.assert_called_once_with(3)

def test_default_listener() -> None:
    event = MagicMock()
    event.job_id = 123
    event.exception = None

    default_listener(event)

    assert event.job_id == 123
    assert event.exception is None


def test_default_listener_with_exception() -> None:
    event = MagicMock()
    event.job_id = 123
    event.exception = Exception("Boom")
    default_listener(event)

    assert isinstance(event.exception, Exception)
    assert str(event.exception) == "Boom"



