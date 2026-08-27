# İnci Akü PPWR Compliance Suite

**Satış sürümü: v1.0.0** — demo değil, üretim ürünü.

**Kilit: 25.08.2026** — açık/koyu tema, tedarikçi PPWR beyanı, müşteri ZIP (yalnızca Technical File + EU DoC PDF), ana sayfa foto şeridi; **Eksik Veri** menüsü kaldırıldı. Arkadaşlar `git pull` ile alır.

Kontrollü PPWR delivery setleri (salt-okunur) + Workspace source of truth + müşteri ZIP teslimatı.

## Yerel üretim (tek port)

```bat
00_START_PPWR_PRODUCTION.cmd
```

- UI + API: http://127.0.0.1:8791  
- OPEN WORD / OPEN PDF → tarayıcı indirme  
- ZIP → indirme  
- Giriş: `admin` / `160616`

Geliştirme (Vite hot reload):

```bat
00_START_PPWR_YAZILIMI.cmd
```

- UI: http://localhost:5173  
- API: http://127.0.0.1:8791/docs  
- Giriş: `admin` / `160616`  
- Takılı port: `00_KILL_STALE_PORTS.cmd`

Masaüstü pencere (Electron):

```bat
00_START_PPWR_DESKTOP.cmd
```

## Render yayın

1. Repo’yu GitHub’a push et  
2. Render → **New Blueprint** → `render.yaml`  
3. Persistent disk `/data` (delivery + workspace)  
4. Frozen delivery ağacını diske kopyala:  
   `/data/delivery/01_STARTER_INDIVIDUAL_DELIVERY_REV00/…`  
5. Deploy → health: `/api/health`

Docker yerel deneme:

```bat
docker build -t inci-ppwr .
docker run -p 10000:10000 -v "D:\ppwr-data:/data" -e PORT=10000 inci-ppwr
```

## Giriş (auth)

Supabase gerekmez. Yerleşik kullanıcı/şifre + JWT oturum:

| Değişken | Varsayılan |
|----------|------------|
| `INCI_PPWR_ADMIN_USER` | `admin` |
| `INCI_PPWR_ADMIN_PASSWORD` | `160616` |
| `INCI_PPWR_JWT_SECRET` | otomatik dosya |
| `INCI_PPWR_AUTH=0` | auth kapat (yalnızca güvenli lokal) |

Kullanıcılar: `workspace/auth/users.json`  
UI: sol menü **Kullanıcılar** (yalnızca admin) — oluştur / şifre sıfırla / aktif-pasif  
API: `POST /api/auth/users`, `…/password`, `…/active`

## Ortam değişkenleri

| Değişken | Açıklama |
|----------|----------|
| `INCI_PPWR_WEB=1` | Web / Render modu (indirme) |
| `INCI_PPWR_DELIVERY_ROOT` | Frozen delivery kökü |
| `INCI_PPWR_WORKSPACE_ROOT` | Yazılabilir workspace |
| `INCI_PPWR_CANDIDATES_ROOT` | Aday paket kökü |
| `INCI_PPWR_VERSION` | Sürüm etiketi |
| `PORT` | Render dinleme portu |

## Önemli

- Frozen delivery setlerine **yazılmaz**.  
- Her müşteri paketi **DOCX + PDF**.  
- PDF: LibreOffice headless (masaüstü + Docker imajında LO yüklü).  
- Web modunda OPEN WORD/PDF ve ZIP → tarayıcı indirme (**İndir** bağlantısı görünür).

## Teslim öncesi smoke

1. `00_START_PPWR_YAZILIMI.cmd` → http://localhost:5173 → `admin` / `160616`  
2. Ana Sayfa KPI’lar dolu  
3. Ürün Arama → OPEN WORD / OPEN PDF indirir  
4. Müşteri Teslimatı → ZIP indir + görünür **ZIP’i indir**  
5. Kullanıcılar (admin) → yeni kullanıcı oluştur  
6. Production: `00_START_PPWR_PRODUCTION.cmd` → http://127.0.0.1:8791  

## Yığın

| Katman | Teknoloji |
|--------|-----------|
| UI | React + Vite + TypeScript |
| API | FastAPI (UI’yi `app/dist` ile servis eder) |
| Desktop | Electron (opsiyonel) |
| Deploy | Docker + Render Blueprint |
