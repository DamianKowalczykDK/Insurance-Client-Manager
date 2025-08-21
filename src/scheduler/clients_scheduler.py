from apscheduler.events import JobExecutionEvent, EVENT_JOB_EXECUTED, EVENT_JOB_ERROR  # type: ignore
from apscheduler.schedulers.background import BackgroundScheduler  # type: ignore
from config import create_client_service
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)


def job_notify_and_cleanup(days_ahead: int = 1, overdue_days: int = 3) -> None:
    logging.info(f"[START] Job started at {datetime.now().isoformat()}")
    client_service = create_client_service()
    client_service.notify_payment_due_in_days(days_ahead=days_ahead)
    remove_clients = client_service.remove_overdue_clients(overdue_days)
    logging.info(f"[DONE] Remove overdue clients {remove_clients}")


def default_listener(event: JobExecutionEvent) -> None:
    if event.exception:
        logging.error(f"[ERROR] Job {event.job_id} failed: {event.exception} ")
    else:
        logging.info(f"[SUCCESS] Job {event.job_id} succeeded")

def create_scheduler(days_ahead: int = 1, overdue_days: int = 3) -> BackgroundScheduler:
    scheduler = BackgroundScheduler()
    scheduler.add_listener(default_listener, EVENT_JOB_EXECUTED | EVENT_JOB_ERROR)

    # scheduler.add_job(
    #     func=job_notify_and_cleanup,
    #     kwargs={"days_ahead": days_ahead, "overdue_days": overdue_days},
    #     trigger="cron",
    #     hour= 1,
    #     minute= 0,
    #     id="job_notify_and_cleanup",
    #     replace_existing=True,
    # )

    scheduler.add_job(
        func=job_notify_and_cleanup,
        kwargs={"days_ahead": days_ahead, "overdue_days": overdue_days},
        trigger="interval",
        seconds=30,
        id="job_notify_and_cleanup",
        replace_existing=True,
    )

    return scheduler