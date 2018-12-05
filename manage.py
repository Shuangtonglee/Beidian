from beidian import app,scheduler,Config,my_listener
from apscheduler.events import EVENT_JOB_EXECUTED,EVENT_JOB_ERROR

if __name__ == '__main__':
    app.config.from_object(Config())
    scheduler.init_app(app)
    scheduler.add_listener(my_listener,EVENT_JOB_EXECUTED | EVENT_JOB_ERROR)
    scheduler.start()
    app.run()