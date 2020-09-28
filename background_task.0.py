import os
import threading
import time

from urllib.request import urlopen

class ImageDownloader(threading.Thread):

    def __init__(self, function_that_downloads):
        threading.Thread.__init__(self)
        self.runnable = function_that_downloads
        self.daemon = True

    def run(self):
        self.runnable()


def downloads():
    with open('somefile.html', 'w+') as f:
        print("Joestar")
        try:
            html = urlopen("http://www.google.er/").read()
            f.write(html)
            print("JOJO")
            print(html)

        except:
            print("DIO")
            f.write('sorry no dice')


print('hi there user')
print('how are you today?')
thread = ImageDownloader(downloads)
thread.start()
while not os.path.exists('somefile.html'):
    print('i am executing but the thread has started to download')
    time.sleep(1)

print('look ma, thread is not alive: ', thread.is_alive())