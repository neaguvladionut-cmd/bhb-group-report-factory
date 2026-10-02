# Group Report Factory

Instrument local, offline, pentru generarea raportului de grup din exporturi Assessment Centre. Primește exportul AC de sinteză, exportul AC detaliat și, opțional, fișierul de declinații; produce auditul XLSX, raportul PowerPoint Trend și bundle-ul de elemente reutilizabile.

## Operator

1. Rulează `npm run build`, apoi deschide `deploy/index.html` în browser.
2. Introdu numele proiectului și încarcă exportul AC de sinteză (scoruri 1–5) și exportul AC detaliat (scoruri 0–2).
3. Rezolvă blocajele și marchează avertismentele după revizie. Zero rămâne zero; celula goală este lipsă, nu zero. Participanții fără niciun scor rămân în audit și sunt excluși din calcule.
4. Configurează grupurile CODE, anexele și secțiunile, apoi descarcă auditul XLSX, raportul Trend și bundle-ul.

## Livrabile

Raportul Trend păstrează cadrele native din template-ul curățat, clonează câte un slide pentru fiecare secțiune din plan și actualizează fiecare grafic împreună cu registrul Excel încorporat care îi aparține. Planul cuprinde coperta, metodologia, imaginea de ansamblu, rezultate de grup, benchmark, observații, comportamente, concluzii și anexele configurate.

Bundle-ul conține câte un SVG cu text convertit în glyph paths, PNG-ul corespunzător, un `data.xlsx`, un manifest și un PDF cu fontul Poppins inclus. Fișierele sunt generate local; nu sunt încărcate în API, bază de date, analytics sau browser storage.

## Date și validare

- AC summary: antete `name`, `cod cp` și competențe cu scor 1–5.
- AC detailed: antete `name the person evaluated`, `cod ac`, `Competente` și `behavior`, cu scoruri 0–2.
- Devplan opțional: `competency`, `behavior`, `objective_text_score_0` și `objective_text_score_2`.

Validarea oprește livrabilul pentru antete necunoscute, identități duplicate, nume sau cod de evaluare lipsă, scor în afara scalei ori benchmark invalid. Identitatea stabilă este persoană + cod evaluare. Lipsa unui scor este avertisment și nu intră în agregare. Grupurile, regiunile, clasamentele și graficele sunt recalculate din participanții proprii fiecărei vederi.

## Dezvoltare

`npm run verify` reconstruiește `deploy/` și rulează testele sintetice. `npm run check:root` confirmă că rădăcina servită este byte-identică cu `deploy/`. Nu sunt folosite date de client în teste, fixture-uri sau commituri.

Dependențele browser sunt livrate local în `src/assets/vendor/`: SheetJS, JSZip, opentype.js și fontul Poppins Regular, cu licențele aferente. Nicio dependență de runtime nu este încărcată din rețea.
