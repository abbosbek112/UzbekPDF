# UzbekPDF — FastAPI + PDF asboblari.
#
# Ilova uchta tashqi dasturni to'g'ridan-to'g'ri chaqiradi (`main.py`):
#   * /usr/bin/gs          — PDF hajmini kichraytirish (Ghostscript)
#   * /usr/bin/libreoffice — Office ↔ PDF konvertatsiya (headless)
#   * poppler (pdf2image)  — PDF sahifalarini rasmga aylantirish
# Shuning uchun ular image ichida bo'lishi shart: usiz ilova ishga tushadi,
# lekin foydalanuvchi tugmani bosgan zahoti xato beradi.

FROM python:3.11-slim

# LibreOffice'ning butuni emas, faqat kerakli ikki qismi: Writer (PDF ga
# aylantirish) va Calc (xlsx). To'liq paket ~700 MB qo'shimcha joy oladi va
# undan hech narsa ishlatilmaydi.
RUN apt-get update && apt-get install -y --no-install-recommends \
        ghostscript \
        poppler-utils \
        libreoffice-writer-nogui \
        libreoffice-calc-nogui \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

# Kod `/usr/bin/libreoffice` ni qat'iy yozib qo'ygan, `-nogui` paketlari esa
# faqat `/usr/bin/soffice` ni beradi. Kodga tegmasdan havola qo'yamiz.
RUN ln -sf /usr/bin/soffice /usr/bin/libreoffice

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY . .

# Vaqtinchalik fayllar va baza volume'da yashaydi — konteyner yangilanganda
# foydalanuvchilar yo'qolib ketmasin.
RUN mkdir -p /app/temp /app/data

ENV PYTHONUNBUFFERED=1

# `-w 2` ataylab: har bir worker LibreOffice va Ghostscript jarayonlarini
# ishga tushiradi, ular esa xotirani ko'p yeydi. Serverda QRdasturxon ham
# turibdi, ya'ni bu yerda ochko'zlik qilib bo'lmaydi.
CMD ["gunicorn", "main:app", "-w", "2", "-k", "uvicorn.workers.UvicornWorker", \
     "-b", "0.0.0.0:8000", "--timeout", "180"]
