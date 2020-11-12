import base64
import uuid
import os
import re
from bs4 import BeautifulSoup as bs
from flask import current_app, url_for

"""import os
import random
import string

path = "D:\\01_test\\DE1078-T-SAG-IMR-00001"


def id_generator(size, chars=string.ascii_uppercase + string.digits):
    return ''.join(random.choice(chars) for _ in range(size))


for path, dirs, files in os.walk(path):
    for folder in dirs:
        os.rename(folder, id_generator(8))
"""


source = '<div class="post-head"><h1 class="post-title">Deploy Flask Applications with uWSGI and Nginx</h1><div class="post-meta"><div class="meta-item author"><a href="/author/todd/"><svg stroke="currentColor" fill="currentColor" stroke-width="0" viewBox="0 0 1024 1024" height="1em" width="1em" xmlns="http://www.w3.org/2000/svg"><path d="M858.5 763.6a374 374 0 0 0-80.6-119.5 375.63 375.63 0 0 0-119.5-80.6c-.4-.2-.8-.3-1.2-.5C719.5 518 760 444.7 760 362c0-137-111-248-248-248S264 225 264 362c0 82.7 40.5 156 102.8 201.1-.4.2-.8.3-1.2.5-44.8 18.9-85 46-119.5 80.6a375.63 375.63 0 0 0-80.6 119.5A371.7 371.7 0 0 0 136 901.8a8 8 0 0 0 8 8.2h60c4.4 0 7.9-3.5 8-7.8 2-77.2 33-149.5 87.8-204.3 56.7-56.7 132-87.9 212.2-87.9s155.5 31.2 212.2 87.9C779 752.7 810 825 812 902.2c.1 4.4 3.6 7.8 8 7.8h60a8 8 0 0 0 8-8.2c-1-47.8-10.9-94.3-29.5-138.2zM512 534c-45.9 0-89.1-17.9-121.6-50.4S340 407.9 340 362c0-45.9 17.9-89.1 50.4-121.6S466.1 190 512 190s89.1 17.9 121.6 50.4S684 316.1 684 362c0 45.9-17.9 89.1-50.4 121.6S557.9 534 512 534z"></path></svg><span>Todd</span></a></div><div class="meta-item tag"><svg stroke="currentColor" fill="currentColor" stroke-width="0" viewBox="0 0 1024 1024" height="1em" width="1em" xmlns="http://www.w3.org/2000/svg"><path d="M483.2 790.3L861.4 412c1.7-1.7 2.5-4 2.3-6.3l-25.5-301.4c-.7-7.8-6.8-13.9-14.6-14.6L522.2 64.3c-2.3-.2-4.7.6-6.3 2.3L137.7 444.8a8.03 8.03 0 0 0 0 11.3l334.2 334.2c3.1 3.2 8.2 3.2 11.3 0zm62.6-651.7l224.6 19 19 224.6L477.5 694 233.9 450.5l311.9-311.9zm60.16 186.23a48 48 0 1 0 67.88-67.89 48 48 0 1 0-67.88 67.89zM889.7 539.8l-39.6-39.5a8.03 8.03 0 0 0-11.3 0l-362 361.3-237.6-237a8.03 8.03 0 0 0-11.3 0l-39.6 39.5a8.03 8.03 0 0 0 0 11.3l243.2 242.8 39.6 39.5c3.1 3.1 8.2 3.1 11.3 0l407.3-406.6c3.1-3.1 3.1-8.2 0-11.3z"></path></svg><span class=""><a class="" href="/tag/devops">DevOps</a></span></div><div class="meta-item reading-time"><svg stroke="currentColor" fill="currentColor" stroke-width="0" viewBox="0 0 1024 1024" height="1em" width="1em" xmlns="http://www.w3.org/2000/svg"><path d="M942.2 486.2C847.4 286.5 704.1 186 512 186c-192.2 0-335.4 100.5-430.2 300.3a60.3 60.3 0 0 0 0 51.5C176.6 737.5 319.9 838 512 838c192.2 0 335.4-100.5 430.2-300.3 7.7-16.2 7.7-35 0-51.5zM512 766c-161.3 0-279.4-81.8-362.7-254C232.6 339.8 350.7 258 512 258c161.3 0 279.4 81.8 362.7 254C791.5 684.2 673.4 766 512 766zm-4-430c-97.2 0-176 78.8-176 176s78.8 176 176 176 176-78.8 176-176-78.8-176-176-176zm0 288c-61.9 0-112-50.1-112-112s50.1-112 112-112 112 50.1 112 112-50.1 112-112 112z"></path></svg><span>10 min read</span></div><div class="meta-item date"><svg stroke="currentColor" fill="currentColor" stroke-width="0" viewBox="0 0 1024 1024" height="1em" width="1em" xmlns="http://www.w3.org/2000/svg"><path d="M880 184H712v-64c0-4.4-3.6-8-8-8h-56c-4.4 0-8 3.6-8 8v64H384v-64c0-4.4-3.6-8-8-8h-56c-4.4 0-8 3.6-8 8v64H144c-17.7 0-32 14.3-32 32v664c0 17.7 14.3 32 32 32h736c17.7 0 32-14.3 32-32V216c0-17.7-14.3-32-32-32zm-40 656H184V460h656v380zM184 392V256h128v48c0 4.4 3.6 8 8 8h56c4.4 0 8-3.6 8-8v-48h256v48c0 4.4 3.6 8 8 8h56c4.4 0 8-3.6 8-8v-48h128v136H184z"></path></svg><span>February 22</span></div></div><figure class="post-image"><img class="post-card-image ls-is-cached lazyloaded" data-src="https://hackersandslackers-cdn.storage.googleapis.com/2020/02/uWSGI@2x.jpg" alt="Deploy Flask Applications with uWSGI and Nginx" src="https://hackersandslackers-cdn.storage.googleapis.com/2020/02/uWSGI@2x.jpg"></figure></div>'
path = "d:\\"


def img_proc(src):
    data = bs(src, "html.parser")
    for img in data.find_all("img"):
        img_src = img["src"]
        if not os.path.exists(os.path.normpath(os.path.join(path, img_src))):
            try:
                filename = str(uuid.uuid4())
                b64string = str(re.findall(r"(?<=base64,)(.*)", img_src))
                extension = str(re.findall(r"(?<=image/)(.*)(?=;base64)", img_src)[0])

                with open(os.path.normpath(os.path.join(path, filename + "." + extension)), "wb") as fh:
                    fh.write(base64.b64decode(b64string))
                img["src"] = os.path.normpath(os.path.join(path, os.path.basename(fh.name)))
                img["data-filename"] = os.path.basename(fh.name)
            except IndexError as err:
                print(f"missing file --> {err}")
        else:
            continue
    return data.prettify()


# a = img_proc(source)
# print(a)


def path_components(path):
    folders = []
    while 1:
        path, folder = os.path.split(path)
        if folder != "":
            folders.append(folder)
        else:
            if path != "":
                folders.append(path)
            break
    folders.reverse()
    return folders


root_path = os.path.normpath("D:/00_herne/ALL_DOCS_download")
exclude = []

for path, dirs, files in os.walk(root_path):
    dirs[:] = [d for d in dirs if d not in exclude]
    for filename in files:
        if "pdf" in filename:
            s = os.path.join(path, filename)
            d = os.path.normpath("\\".join(path_components(os.path.join(path, filename))[4:]))
            print(d)

