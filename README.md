In terminal, navigate to the folder called rrd_ocr

Place the pdfs you want to convert to docx format in the folder called src/Dataset/

RUN the following commands:

	sudo docker build -t  rrd_ocr .

	sudo docker run -it --mount src="$(pwd)/src",target=/src,type=bind rrd_ocr

Now let the excecution process complete. 

When it is done, all the docx files are available in the folder called finale.




