1) Codefile_android.gs را جایگزین کد Apps Script کنید (فقط doPost اضافه شده) و Deploy > New version کنید (Web app، Anyone).
2) در www/index.html مقدار API_URL را با آدرس /exec پر کنید.
3) npm install && npx cap add android && npx cap sync && cd android && ./gradlew assembleDebug
   (یا پوشه را در GitHub بگذارید و workflow «Build APK» را اجرا کنید.)
