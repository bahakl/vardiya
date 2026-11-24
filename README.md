# Vardiya

Kamera ekibi vardiya uygulaması (Türkçe). Statik site (GitHub Pages) + Google Apps Script arka ucu.

- `index.html`, `config.js`: site. `config.js` içindeki `VARDIYA_API`, Apps Script web uygulamasının `/exec` adresidir.
- `apps-script/Code.gs`: arka uç. Çizelge, telefon listeleri ve uygulama verisi Google E-Tablonuzda durur; bu depoda veri yoktur.

Kurulum: Code.gs'i bir Apps Script projesine yapıştırın, `SHEET_ID` değerini e-tablo kimliğinizle değiştirin, `setup()` çalıştırın, Web uygulaması olarak dağıtın (Me / Anyone), adresi `config.js`'e yazın.
