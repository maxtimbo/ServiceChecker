import psutil

def get_service(name):
    service = None
    try:
        service = psutil.win_service_get(name)
        service = service.as_dict()
    except Exception as ex:
        pass
        # print(ex)
    return service

def check_service(name):
    service = get_service(name)
    if service:
        print("Service found: ", service)
        if service and service['status'] == 'running':
            print("Service is running")
        else:
            print("Service is not running")
    else:
        print("Service not found")

def list_services():
    x = 0
    for serv in psutil.win_service_iter():
        x += 1
        try:
            servname = serv.name()
            servdisp = serv.display_name()
            servstatus = serv.status()
        except Exception as ex:
            print("reaised exc", ex)
            pass

list_services()
