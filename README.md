# Okul Tanıtım Randevu Sistemi - CORS Düzeltilmiş Sürüm

Bu sürüm GitHub Pages + Google Apps Script arasında oluşan CORS engelini JSONP yöntemiyle çözer.

Kurulum:
1. Google Sheets > Uzantılar > Apps Script içine AppsScript_OkulTanitim_CORS_Duzeltilmis.gs kodunu yapıştırın.
2. setupWorkbook fonksiyonunu bir kez çalıştırın.
3. Web app olarak yeniden deploy edin.
4. Web app URL adresini index.html içindeki SCRIPT_URL alanına yazın.
5. index.html ve 404.html dosyalarını GitHub reposuna yükleyin.
