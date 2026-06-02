# İngiliz Kültür Koleji - Okul Tanıtım Randevu Sistemi

Bu sürümde veli e-posta alanı ve otomatik HTML randevu bileti maili vardır.

## Kurulum

1. `OkulTanitimRandevu_TabloSablonu_VeliMail.xlsx` dosyasını Google Drive'a yükleyin ve Google Sheets olarak açın.
2. Google Sheets > Uzantılar > Apps Script bölümünü açın.
3. `AppsScript_OkulTanitim_VeliMail.gs` dosyasındaki kodu Apps Script'e yapıştırın.
4. `setupWorkbook()` fonksiyonunu bir kez çalıştırın.
5. Deploy > New deployment > Web app seçin.
   - Execute as: Me
   - Who has access: Anyone
6. Web App URL'sini alın.
7. `index.html` içindeki `SCRIPT_URL` alanına bu URL'yi yapıştırın.
8. `index.html` ve `404.html` dosyalarını GitHub Pages reposuna yükleyin.

## Admin Panel

Admin adresi:

```text
siteadresiniz.com/admin
```

İlk PIN: `1234`

PIN'i Google Sheets > Ayarlar sayfasındaki `admin_pin` satırından değiştirin.

## Mail Sistemi

Her randevu şu adreslere mail düşer:

- info@ingilizkultur.com.tr
- batikent@ingilizkultur.com.tr

Veli e-posta yazdığı için ayrıca veliye İngiliz tren/peron bileti tarzında HTML randevu bileti gönderilir.

Veli mailini kapatmak isterseniz Google Sheets > Ayarlar sayfasında:

```text
parent_mail_enabled = FALSE
```

olarak değiştirin.
