FROM python:3.9
WORKDIR /
COPY requirements.txt . 
RUN apt-get --allow-releaseinfo-change update && apt-get update && apt-get install -y libgl1
RUN pip install -r requirements.txt
COPY / .
CMD ["python", "./main.py"]