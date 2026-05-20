# MRP-Analysesystem — Industrielles Lasermodul

> 🇩🇪 Deutsch | [🇬🇧 English](#english) | [🇹🇷 Türkçe](#türkçe)

---

Ein modulares **Material Requirements Planning (MRP)** System, entwickelt mit Python und Excel.
Als Beispielszenario dient ein industrielles Lasermodul — angelehnt an Produkte von Hightech-Unternehmen wie Jenoptik, Carl Zeiss oder SCHOTT.
Das System ist flexibel gestaltet und kann leicht auf andere Produkte und Branchen angepasst werden.

---

## Funktionen

**Material Requirements Planning (MRP)**
- Mehrstufige Stückliste (BOM) — 3 Ebenen: Endprodukt, Baugruppe, Rohteil
- Wöchentliche Brutto- und Nettobedarfsberechnung
- Dynamische Bestandsführung unter Berücksichtigung des Mindestbestands
- Bestellwochenberechnung auf Basis der Durchlaufzeit

**Kritikalitätsanalyse**
- 3 Risikofaktoren: Kritikalitätsrisiko, Lieferrisiko, Haltbarkeitsrisiko
- Bewertung auf einer Skala von 1–10 — hoher Wert = hohes Risiko
- Faktoren werden per Multiplikation kombiniert (in Anlehnung an FMEA-Methodik)
- Mindestbestand wird mit CEILING berechnet — immer auf der sicheren Seite

**Kapazitätsplanung**
- Wöchentliche Taktzeit- und Zykluszeit-Analyse
- Unterstützung von Effizienzfaktor und Wartungswochen
- Farbkodierter Kapazitätsstatus: Idealbereich / Achtung / Engpass / Unterauslastung
- Zusammenfassende Kapazitätsstatistiken

---

## Warum Multiplikation statt gewichtetem Mittel?

In der Hightech-Fertigung bestimmt das schwächste Glied das gesamte System.
Fehlt die Lasertreiberplatine, stoppt die Produktion — eine lange Haltbarkeit kann das nicht ausgleichen.
Das gewichtete Mittel verbirgt dieses Risiko, die Multiplikation nicht.

Diese Methodik ist identisch mit der FMEA-RPZ-Berechnung:

```
RPZ = Bedeutung × Auftretenswahrscheinlichkeit × Entdeckungswahrscheinlichkeit
```

---

## Projektstruktur

```
mrp-rechner/
│
├── kritiklik.py       # Kritikalitätsfaktor & Mindestbestandsberechnung
├── mrp.py             # MRP-Hauptberechnungsmodul
├── kapasite.py        # Kapazitätsplanung & Taktzeit/Zykluszeit
├── main.py            # Hauptprogramm
│
├── mrp_project.xlsx   # Eingabedaten (5 Arbeitsblätter)
│   ├── Stückliste          → Mehrstufige Materialliste
│   ├── Lagerbestand        → Bestandsübersicht & Kritikalitätsanalyse
│   ├── Produktionsplan     → 12-wöchiger Produktionsplan (MPS)
│   ├── Kapazitätsplanung   → Wöchentliche Kapazitätsanalyse
│   └── Bewertungsmodell    → Erläuterung des Bewertungssystems
│
└── mrp_ergebnis.xlsx  # Ausgabe (wird automatisch erstellt)
    ├── Lageranalyse        → Kritikalitäts- & Bestandsanalyseergebnisse
    ├── MRP_Ergebnis        → Brutto-/Nettobedarf & Bestellwochen
    ├── Bestellvorschlag    → Nur Positionen mit Bestellbedarf
    └── Kapazitätsanalyse   → Wöchentliche Kapazitätsauslastung
```

---

## Installation & Ausführung

```bash
pip install pandas openpyxl
python main.py
```

---

## Verwendete Bibliotheken

| Bibliothek | Verwendungszweck |
|---|---|
| `pandas` | Datenverarbeitung und Excel-Import/-Export |
| `openpyxl` | Excel-Formatierung |
| `math` | CEILING-Funktion für ganzzahlige Rundung |

---

## Schlüsselbegriffe

| Begriff | Erklärung |
|---|---|
| MRP | Material Requirements Planning — wann und wie viel bestellen? |
| BOM | Bill of Materials — welche Teile in welcher Menge pro Produkt? |
| Taktzeit | Maximal verfügbare Zeit pro Einheit zur Deckung der Kundennachfrage |
| Zykluszeit | Tatsächliche Produktionszeit pro Einheit |
| Kritikalitätsfaktor | Multiplikation der 3 Risikofaktoren — angelehnt an FMEA |
| Mindestbestand | Sicherheitspuffer für unterbrechungsfreie Produktion |

---

## Excel-Datei — Hinweise

**Gelbe Zellen** = Eingabefelder (können geändert werden): Bedarfsmengen, Risikobewertungen, Durchlaufzeiten.
**Graue kursive Zellen** = Formeln (nicht verändern).

---

## Beispielausgabe

```
============================================================
  Jenoptik – Lasermodul | MRP-Analysesystem
============================================================
Daten eingelesen:
  Stückliste      : 29 Einträge
  Lagerbestand    : 24 Bauteile
  Produktionsplan : 12 Wochen

-- Kritikalitäts- & Lageranalyse --
Gesamt: 24 | Ausreichend: 23 | Kritisch: 1

⚠ KRITISCHER BESTAND — SOFORTMASSNAHME ERFORDERLICH:
  [DK-212] Lasertreiberplatine
  Aktuell: 6 | Mindestbestand: 7

-- Kapazitätszusammenfassung --
  Durchschnittliche Auslastung : %111.5
  Engpass-Wochen               : 8
  Wochen im Idealbereich       : 3
============================================================
  Analyse abgeschlossen.
============================================================
```

---

## Erweiterungsmöglichkeiten

- Mehrstufige BOM-Auflösung (derzeit einsstufige MRP-Berechnung)
- Szenarioanalyse — Vergleich bei veränderten Zykluszeiten oder Effizienzwerten
- Visualisierung — Gantt-Diagramm für Produktionsplanung
- API-Schnittstelle für SAP/ERP-Integration

---
---

# English

> [🇩🇪 Deutsch](#mrp-analysesystem--industrielles-lasermodul) | 🇬🇧 English | [🇹🇷 Türkçe](#türkçe)

---

A modular **Material Requirements Planning (MRP)** system built with Python and Excel.
The example scenario is based on an industrial laser module — inspired by products from Hightech companies such as Jenoptik, Carl Zeiss, or SCHOTT.
The system is flexible and can easily be adapted to other products and industries.

---

## Features

**Material Requirements Planning (MRP)**
- Multi-level Bill of Materials (BOM) — 3 levels: End product, sub-assembly, raw part
- Weekly gross and net requirements calculation
- Dynamic inventory update with safety stock reservation
- Order week calculation based on lead time

**Criticality Analysis**
- 3 risk factors: Criticality risk, supply risk, shelf-life risk
- Scored on a scale of 1–10 — higher score = higher risk
- Factors combined via multiplication (inspired by FMEA methodology)
- Safety stock calculated with CEILING — always on the safe side

**Capacity Planning**
- Weekly Takt Time and Cycle Time analysis
- Efficiency factor and maintenance week support
- Color-coded capacity status: Ideal / Caution / Bottleneck / Underutilized
- Summary capacity statistics

---

## Why Multiplication Instead of Weighted Average?

In high-tech manufacturing, the weakest link determines the entire system.
If the laser driver board is missing, production stops — a long shelf life cannot compensate for that.
A weighted average hides this risk; multiplication does not.

This approach is identical to the FMEA Risk Priority Number (RPN) calculation:

```
RPN = Severity × Occurrence × Detection
```

---

## Project Structure

```
mrp-rechner/
│
├── kritiklik.py       # Criticality factor & safety stock calculation
├── mrp.py             # MRP main calculation module
├── kapasite.py        # Capacity planning & Takt/Cycle Time
├── main.py            # Main execution file
│
├── mrp_project.xlsx   # Input data (5 worksheets)
└── mrp_ergebnis.xlsx  # Output (auto-generated)
```

---

## Installation & Usage

```bash
pip install pandas openpyxl
python main.py
```

---

## Key Concepts

| Term | Description |
|---|---|
| MRP | Material Requirements Planning — when and how much to order? |
| BOM | Bill of Materials — which parts, in what quantity per product? |
| Takt Time | Maximum time per unit to meet customer demand |
| Cycle Time | Actual production time per unit |
| Criticality Factor | Multiplication of 3 risk factors — inspired by FMEA |
| Safety Stock | Buffer stock for uninterrupted production |

---

## Development Ideas

- Multi-level BOM explosion (currently single-level MRP)
- Scenario analysis — compare results with different cycle times or efficiency values
- Visualization — Gantt chart for production scheduling
- API layer for SAP/ERP integration

---
---

# Türkçe

> [🇩🇪 Deutsch](#mrp-analysesystem--industrielles-lasermodul) | [🇬🇧 English](#english) | 🇹🇷 Türkçe

---

Python ve Excel ile geliştirilmiş modüler bir **Malzeme İhtiyaç Planlaması (MRP)** sistemi.
Örnek senaryo olarak endüstriyel bir lazer modülü kullanılmıştır — Jenoptik, Carl Zeiss veya SCHOTT gibi yüksek teknoloji şirketlerinin ürünlerinden ilham alınmıştır.
Sistem esnek bir yapıya sahiptir ve farklı ürün ve sektörlere kolayca uyarlanabilir.

---

## Özellikler

**Malzeme İhtiyaç Planlaması (MRP)**
- Çok seviyeli Malzeme Listesi (BOM) — 3 seviye: Son ürün, alt montaj, ham parça
- Haftalık brüt ve net ihtiyaç hesabı
- Minimum stok rezervi korunarak dinamik stok güncellemesi
- Sipariş haftası hesabı (temin süresi baz alınarak)

**Kritiklik Analizi**
- 3 risk faktörü: Kritiklik riski, tedarik riski, raf ömrü riski
- Her faktör 1-10 arası puanlanır — yüksek puan = yüksek risk
- Katsayılar çarpım yöntemiyle birleştirilir (FMEA metodolojisi)
- Minimum stok CEILING ile hesaplanır — güvenli tarafta kal

**Kapasite Planlaması**
- Haftalık Takt Time ve Cycle Time analizi
- Verimlilik katsayısı ve bakım haftası desteği
- Renk kodlu kapasite durumu: İdeal / Dikkat / Darboğaz / Atıl kapasite
- Kapasite özet istatistikleri

---

## Neden Çarpım, Neden Ağırlıklı Ortalama Değil?

Yüksek teknoloji üretiminde en zayıf halka tüm sistemi belirler.
Lazer Sürücü Devre Kartı gelmezse fabrika durur — raf ömrünün uzun olması bunu telafi edemez.
Ağırlıklı ortalama bu riski gizler, çarpım gizleyemez.

Bu yaklaşım FMEA metodolojisindeki RPN hesabıyla özdeştir:

```
RPN = Severity × Occurrence × Detection
```

---

## Kurulum & Çalıştırma

```bash
pip install pandas openpyxl
python main.py
```

---

## Temel Kavramlar

| Kavram | Açıklama |
|---|---|
| MRP | Malzeme İhtiyaç Planlaması — ne zaman, ne kadar sipariş verilmeli? |
| BOM | Malzeme Listesi — 1 ürün için hangi parçalar, kaç adet? |
| Takt Time | Müşteri talebini karşılamak için birim başına max. süre |
| Cycle Time | Gerçekte bir ürünü üretmek için harcanan süre |
| Kritiklik Katsayısı | 3 risk faktörünün çarpımı — FMEA'dan ilham alınmıştır |
| Minimum Stok | Üretim güvenliği için korunması gereken tampon stok |

---

## Excel Dosyası Kullanımı

**Sarı hücreler** = Elle girilebilir (talep miktarları, risk puanları, temin süreleri).
**Gri italik hücreler** = Formülle hesaplanır, dokunmayınız.

---

## Geliştirme Fikirleri

- Çok seviyeli BOM hesaplaması
- Senaryo analizi — farklı cycle time ve verimlilik değerleriyle karşılaştırma
- Görselleştirme — Gantt chart ile üretim çizelgesi
- SAP/ERP entegrasyonu için API katmanı
