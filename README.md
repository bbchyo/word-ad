# EBYÜ Tez Format Doğrulayıcı - Word Add-in

Erzincan Binali Yıldırım Üniversitesi Tez Yazım Kılavuzu'na göre Word belgelerini kontrol eden bir Office Add-in.

![EBYÜ Thesis Validator](https://img.shields.io/badge/EBYÜ-Thesis%20Validator-blue)
![Office.js](https://img.shields.io/badge/Office.js-1.1+-green)
![License](https://img.shields.io/badge/license-MIT-blue)

## 📁 Proje Yapısı

```
word-ad/
├── index.html          # Ana UI dosyası
├── taskpane.css        # Stil dosyası  
├── taskpane.js         # Office.js mantığı
├── manifest.xml        # Add-in yapılandırması
├── vercel.json         # Vercel deployment config
├── assets/
│   ├── icon-16.svg
│   ├── icon-32.svg
│   ├── icon-64.svg
│   └── icon-80.svg
└── README.md
```

## ✅ Kontrol Edilen Kurallar (EBYÜ 2022 Kılavuzu)

### 1. Kenar Boşlukları
- Tüm kenarlar: **3 cm** (85 point)

### 2. Yazı Tipi
- Genel metin: **Times New Roman, 12 pt**
- Dipnotlar: **Times New Roman, 10 pt**
- Tablo içi metin: **Times New Roman, 11 pt**
- Ana başlıklar: **14 pt, Kalın, BÜYÜK HARF, Ortalı**
- Alt başlıklar: **12 pt, Kalın, 1.25 cm girinti**

### 3. Paragraf Formatı
- Hizalama: **İki yana yaslı (Justify)**
- Satır aralığı: **1.5 satır**
- İlk satır girintisi: **1.25 cm**
- Paragraf aralığı: **Önce 6pt, Sonra 6pt**

### 4. Görseller
- Sarmalama: **Metinle aynı hizada** (kayma önleme)
- Hizalama: **Ortalanmış**

### 5. Tablolar (YENİ!)
- **Gizli/Düzen Tabloları**: Kenarlıksız tablolar tespit edilir, "Sütunlar" kullanımı önerilir
- **Başlık Konumu**: Tablo başlığı tablonun **ÜSTünde** olmalıdır
- **Yazı Boyutu**: Tablo içi metin **11 pt** olmalıdır
- **Satır Aralığı**: Tablolarda **tek (1.0)** satır aralığı kullanılmalıdır
- **Genişlik**: Tablo sayfa kenar boşluklarını aşmamalıdır

## 🚀 Vercel'e Dağıtım

### 1. GitHub'a Yükle

```bash
cd word-ad
git init
git add .
git commit -m "EBYÜ Thesis Validator - Word Add-in"
git branch -M main
git remote add origin https://github.com/KULLANICI/word-ad.git
git push -u origin main
```

### 2. Vercel'de Dağıt

1. [Vercel Dashboard](https://vercel.com/dashboard)'a git
2. "Add New" → "Project" tıkla
3. GitHub reposunu seç
4. "Deploy" tıkla
5. Deployment URL'ini kopyala (örn: `https://word-ad.vercel.app`)

### 3. Manifest URL'lerini Güncelle

`manifest.xml` dosyasındaki tüm `https://localhost:3000` URL'lerini Vercel URL'iniz ile değiştirin:

```bash
# macOS/Linux
sed -i '' 's|https://localhost:3000|https://YOUR-APP.vercel.app|g' manifest.xml

# veya manuel olarak düzenleyin
```

## 📱 Add-in'i Word'e Yükleme

### Yöntem 1: Sideloading (Test için)

**Windows:**
1. Word'ü aç
2. `Insert` → `My Add-ins` → `Upload My Add-in`
3. `manifest.xml` dosyasını seç

**Mac:**
1. Word'ü aç
2. `Insert` → `Add-ins` → `My Add-ins`
3. Sol alt köşede `Upload My Add-in`
4. `manifest.xml` dosyasını seç

### Yöntem 2: SharePoint Catalog (Kurumsal)

1. SharePoint'te bir App Catalog oluşturun
2. `manifest.xml` dosyasını yükleyin
3. Kullanıcılar Add-in'i Word'den ekleyebilir

## 🔧 Yerel Geliştirme

```bash
# Yerel sunucu başlat
npx http-server -p 3000 --cors

# Tarayıcıda aç
open http://localhost:3000
```

## 🛠️ Özellikler

- ✅ **Otomatik Tarama**: Tek tıkla tüm belgeyi kontrol et
- ✅ **Hata Kategorileri**: Kırmızı (Hata), Sarı (Uyarı), Yeşil (Başarılı)
- ✅ **Otomatik Düzeltme**: Yazı tipi, kenar boşlukları ve satır aralığını tek tıkla düzelt
- ✅ **Tablo Kontrolü**: Gizli tablolar, başlık konumu, yazı boyutu kontrolü
- ✅ **Türkçe Arayüz**: Tamamen Türkçe kullanıcı deneyimi

## 📋 API Gereksinimleri

- Office.js 1.1+
- Word 2016+ veya Microsoft 365
- Word Online desteklenir

## 🐛 Sorun Giderme

### Add-in yüklenmiyor
- Manifest URL'lerinin HTTPS olduğundan emin olun
- Word'ü yeniden başlatın
- Cache'i temizleyin: `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`

### Tarama çalışmıyor
- Belgenin boş olmadığından emin olun
- DevTools konsolunu kontrol edin (F12)

## 📄 Lisans

MIT License - Erzincan Binali Yıldırım Üniversitesi

---

**Geliştirici**: EBYÜ Thesis Validator Team  
**Versiyon**: 1.0.0  
**Kılavuz**: EBYÜ 2022 Tez Yazım Kılavuzu
