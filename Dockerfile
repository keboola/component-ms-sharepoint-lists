FROM python:3.7.2-slim
ENV PYTHONIOENCODING utf-8

COPY . /code/

# Debian 9 (stretch), the base of this image, is end of life and its packages
# were moved to archive.debian.org, so "apt-get update" fails with 404 and the
# image cannot be built at all. Nothing in requirements.txt needs a compiler,
# so build-essential is no longer installed.

RUN pip install flake8

RUN pip install -r /code/requirements.txt

WORKDIR /code/


CMD ["python", "-u", "/code/src/component.py"]
