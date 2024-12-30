
---

### **Bu Kodun Özellikleri**
1. **AF'nin Sağına Sütun Ekler:**
   - `Unique Count` ve `İlişki` başlıkları AF'nin hemen sağında eklenir.
   - Bu başlıkların altında hesaplamalar doğru şekilde yapılır.

2. **Unique Count Hesaplaması:**
   - AE sütunundaki değerlerin mağaza bazında tekrar sayısı doğru şekilde hesaplanır.

3. **İlişki Sütunu:**
   - AF sütunundaki değerlere göre doğru şekilde doldurulur: `"Muadil"`, `"Muadil stoksuz"`, `"İlişki yok"`.

4. **Tekli ve Çift Sayfaları:**
   - Unique Count değerine göre filtrelenir ve özel formüller doğru şekilde uygulanır.

---

### **Nasıl Çalıştırılır**
1. Yukarıdaki kodu bir `.py` dosyasına kaydedin (örneğin, `app.py`).
2. Streamlit uygulamanızı başlatmak için:
   ```bash
   streamlit run app.py
