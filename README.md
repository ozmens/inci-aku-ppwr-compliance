# İnci Akü PPWR Compliance Suite

v1.0.0 — İnci Akü’ye özel üretim ürünü. Tüm kontrollü evraklar (Word + PDF) programın parçasıdır.

Kilit 25.08.2026: güncel arayüz ve ZIP davranışı GitHub `main` üzerindedir (`git pull`).

## Giriş

`admin` / `160616`

## Yerel / İnci sunucusu

Önce teslimat setlerini bağlayın (bir kez):

```bat
00_LINK_DELIVERY.cmd
```

Sonra:

```bat
00_START_PPWR_PRODUCTION.cmd
```

Adres: http://127.0.0.1:8791

## Render (internet)

1. Bu private repo → Render **New Blueprint** (`render.yaml`)
2. Plan: Standard, disk: 50 GB (`/data`)
3. Teslimat ağacını diske koyun (`00_COPY_DELIVERY.cmd` çıktısı → `/data/delivery`)
4. Sağlık: `/api/health`

İnci ekibi bundan sonraki revizyonları bu ürün üzerinde yapar. Frozen teslimat setlerine yazılmaz.

## Teslim kuralları

- Her müşteri paketi **Word + PDF**
- PDF: LibreOffice (Docker imajında yüklü)
- Web’de OPEN WORD / OPEN PDF ve ZIP tarayıcıya iner
