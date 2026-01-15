# PartyQuiz
Automatické generování kvízových prezentací z Google Tabulek do Google Slides

## Nastavení - krok za krokem

### 1. Google Cloud Console - Vytvoření projektu a povolení API

#### a) Vytvoření projektu (pokud ještě nemáte)

1. Jděte na [Google Cloud Console](https://console.cloud.google.com/)
2. Klikněte na **"Select a project"** (nahoře) → **"New Project"**
3. Zadejte název projektu (např. "PartyQuiz")
4. Klikněte **"Create"**

#### b) Povolení potřebných API

1. V Google Cloud Console vyberte váš projekt
2. Klikněte na **"APIs & Services"** → **"Library"** (v levém menu)
3. Povolte tyto tři API (vyhledejte je a u každého klikněte **"Enable"**):
   - **Google Sheets API**
   - **Google Slides API**
   - **Google Drive API**

#### c) Vytvoření OAuth2 přihlašovacích údajů

1. Jděte na **"APIs & Services"** → **"Credentials"**
2. Klikněte **"Create Credentials"** → **"OAuth client ID"**
3. Pokud se zobrazí upozornění o consent screen:
   - Klikněte **"Configure Consent Screen"**
   - Vyberte **"External"** → **"Create"**
   - Vyplňte pouze povinná pole:
     - App name: např. "PartyQuiz"
     - User support email: váš email
     - Developer contact: váš email
   - Klikněte **"Save and Continue"** třikrát (přeskočte Scopes a Test users)
   - Klikněte **"Back to Dashboard"**
4. Nyní vytvořte credentials:
   - **"Create Credentials"** → **"OAuth client ID"**
   - Application type: **"Desktop app"**
   - Name: např. "PartyQuiz Desktop"
   - Klikněte **"Create"**
5. **Stáhněte JSON soubor** - klikněte na ikonu stahování
6. Přejmenujte stažený soubor na `credentials.json`
7. Přesuňte `credentials.json` do složky s tímto projektem

#### d) Přidání testovacích uživatelů

1. V Google Cloud Console jděte na **"APIs & Services"** → **"OAuth consent screen"**
2. Najděte sekci **"Test users"**
3. Klikněte **"+ ADD USERS"**
4. Zadejte svůj Google email (ten, kterým se budete přihlašovat)
5. Klikněte **"Save"**

**⚠️ DŮLEŽITÉ:** Bez přidání sebe jako testovacího uživatele dostanete chybu "Error 403: access_denied"

### 2. Instalace Python balíčků

V této složce spusťte příkaz:

```bash
pip install -r requirements.txt
```

### 3. Konfigurace skriptu

Otevřete soubor `generate_google_slides.py` a upravte tyto řádky:

```python
SPREADSHEET_ID = "VÁŠ_ID_TABULKY"  # ID z URL tabulky
RANGE_NAME = 'otazky!A:C'  # Název listu a rozsah sloupců
```

**Jak získat ID tabulky:**
- Otevřete tabulku v prohlížeči
- URL vypadá: `https://docs.google.com/spreadsheets/d/19o12YCNMoeEY4UBvScqW2vDv35rbC21RelnUcn__ExM/edit`
- ID je prostřední část: `19o12YCNMoeEY4UBvScqW2vDv35rbC21RelnUcn__ExM`

### 4. Spuštění skriptu

```bash
python generate_google_slides.py
```

**Při prvním spuštění:**
1. Otevře se prohlížeč s přihlášením do Google účtu
2. Přihlaste se účtem, který jste přidali jako testovacího uživatele
3. Pokud se zobrazí varování "Google hasn't verified this app":
   - Klikněte **"Advanced"** nebo **"Pokročilé"**
   - Klikněte **"Go to PartyQuiz (unsafe)"** nebo **"Přejít na PartyQuiz"**
4. Klikněte **"Allow"** nebo **"Povolit"** pro všechna oprávnění
5. Skript vytvoří soubor `token.json` - příště už se přihlašovat nebudete muset

**Další spuštění:**
- Již se nemusíte přihlašovat, skript použije uložený `token.json`

## Jak skript funguje

### Struktura tabulky

Skript očekává tuto strukturu v Google Tabulce:

- **1. řádek:** Záhlaví (přeskakuje se)
- **Prázdný řádek** = oddělovač témat
- **Řádek s tématem:** Název v prvním sloupci (např. "1. Aktuality")
- **6 řádků s otázkami:** 
  - Sloupec A (Téma): Otázka s číslem na začátku (např. "2) Jaká je hlavní města?")
  - Sloupec B (Odpovědi): Odpověď
  - Sloupec C (Poznámky): Poznámky

### Co skript vytvoří

- **1 prezentace na každá 2 témata** (témata 1+2, 3+4, 5+6, atd.)
- **1 slide na každou otázku**
- **Titulek slide:** Číslo otázky + název tématu (např. "2) Aktuality")
- **Tělo slide:** Otázka bez čísla
- Prezentace se automaticky uloží do **stejné složky jako tabulka**

### Příklad výstupu

Pokud máte témata:
- 1. Aktuality (6 otázek)
- 2. Hadi (6 otázek)
- 3. Rok 2025 (6 otázek)

Vytvoří se:
- **Prezentace 1:** "Party Quiz - 1. Aktuality & 2. Hadi" (12 slides)
- **Prezentace 2:** "Party Quiz - 3. Rok 2025" (6 slides)

## Řešení problémů

### "Error 403: access_denied"
- Přidejte svůj email jako testovacího uživatele (viz bod 1d výše)
- Smažte `token.json` a spusťte skript znovu

### "Google Drive API has not been used"
- Povolte Google Drive API v Google Cloud Console (viz bod 1b)
- Počkejte 1-2 minuty a zkuste znovu

### "invalid_grant" nebo zastaralý token
- Smažte `token.json`
- Spusťte skript znovu - otevře se přihlášení

### Prezentace se neuloží do sdílené složky
- Ujistěte se, že je povolené Google Drive API
- Zkontrolujte, že `token.json` byl vytvořen PO povolení Drive API
- Pokud ne, smažte `token.json` a přihlaste se znovu

### Skript nenašel žádná témata
- Zkontrolujte, že tabulka má správnou strukturu
- Ujistěte se, že jsou témata oddělená prázdnými řádky
- Zkontrolujte nastavení `RANGE_NAME` v souboru

## Soubory v projektu

- `generate_google_slides.py` - Hlavní skript
- `credentials.json` - OAuth2 přihlašovací údaje (NESDÍLEJTE!)
- `token.json` - Token pro přihlášení (vytváří se automaticky, NESDÍLEJTE!)
- `requirements.txt` - Seznam potřebných Python balíčků
- `.gitignore` - Chrání citlivé soubory před nahráním na Git

## Poznámky

- Každé spuštění vytvoří **nové prezentace** - nepřepisuje staré
- Pokud chcete změnit tabulku, upravte `SPREADSHEET_ID` v souboru
- Skript podporuje čísla i písmeno "T" na začátku otázek
- Pokud otázka nemá číslo, použije se jen název tématu jako titulek
