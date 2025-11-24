# Bygga Windows Executable med GitHub Actions (Gratis!)

## Steg:

1. **Skapa GitHub-repo** (om du inte har det):
   - Gå till github.com
   - Skapa nytt repository
   - Pusha din kod dit

2. **Workflow är redan skapad** (finns i `.github/workflows/build-windows.yml`)

3. **Kör workflow:**
   - Gå till ditt repo på GitHub
   - Klicka på "Actions"
   - Välj "Build Windows Executable"
   - Klicka "Run workflow"

4. **Ladda ner executable:**
   - När bygget är klart, gå till "Artifacts"
   - Ladda ner `outlook-attach-server-windows`
   - Det innehåller `.exe`-filen!

5. **Lägg i paketet:**
   - Kopiera `.exe` till `server/`-mappen
   - Skapa nytt ZIP-paket

**Fördelar:**
- ✅ Gratis
- ✅ Körs i molnet på riktig Windows
- ✅ Inget att installera på din Mac
- ✅ Fungerar perfekt

