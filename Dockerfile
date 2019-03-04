FROM ubuntu:18.04


RUN apt-get update \
    && apt-get install tesseract-ocr  -y


RUN apt-get install imagemagick -y

RUN apt-get install python3-pip -y

COPY requirements.txt /
RUN pip3 install -r /requirements.txt 

RUN apt-get install python3

RUN pip3 install opencv-python
RUN apt-get install -y libsm6 libxext6


WORKDIR ./src/

COPY ./policy.xml /

COPY ./policy.xml /etc/ImageMagick-6/


CMD [ "python3", "./ocrprog.py" ]
