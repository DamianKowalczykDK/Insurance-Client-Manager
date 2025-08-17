from src.scheduler.clients_scheduler import create_scheduler
import time


def main() -> None:
    scheduler = create_scheduler()
    scheduler.start()
    print(f"[SCHEDULER] starting ... ")

    try:
        while True:
            time.sleep(1)
    except (KeyboardInterrupt, SystemExit):
        scheduler.shutdown()
        print(f"[SCHEDULER] Stopping ... ")

if __name__ == '__main__':
    main()