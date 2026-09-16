# Group Report Factory

Instrument local, offline, pentru raportul de grup din Assessment Centre. Primește exporturile standard deja produse de ecosistemul legacy și descarcă două livrabile editabile: auditul XLSX și raportul PPTX în română.

## Operator

1. Deschide `deploy/index.html` direct în browser sau rulează `npm run build`, apoi deschide fișierul rezultat.
2. Pasul 1 — **Încarcă**: adaugă exportul AC de sinteză (scoruri 1–5) și, pentru comportamente, exportul AC detaliat (scoruri 0–2). Poți selecta fișierele împreună sau prin drag-and-drop.
3. Pasul 2 — **Revizuiește**: parcurge cele șapte sub-etape, rezolvă blocajele și marchează avertismentele după ce le-ai revizuit. Zero observat rămâne zero; o celulă goală este avertisment, nu este transformată în zero. Participanții complet neevaluati rămân vizibili în audit, dar sunt excluși din calcule.
4. Pasul 3 — **Generează**: alege rendererul **BHB** sau **TREND**, controlează secțiunile și livrarea într-un singur PPTX sau în două fișiere separate, apoi descarcă auditul XLSX și livrabilul/livrabilele PowerPoint.

## Raportul generat

Raportul principal include coperta cu marcaj **CONFIDENȚIAL**, metodologia, Executive Summary, rezultatele de grup, benchmarkul, câte un Key Findings pentru fiecare competență, comportamentele relevante și concluziile introduse explicit. Anexa păstrează distribuțiile detaliate și rezultatele individuale. Structura este aceeași pentru ambele renderere; BHB folosește identitatea BHB, iar TREND automatizează cadrele vizuale din `Template - raport de grup - RO.pptx`. În livrarea într-un singur fișier, raportul principal și anexa sunt secțiuni native separate; în livrarea divizată sunt generate separat.

Fiecare grafic implicit este un grafic PowerPoint nativ cu registru Excel încorporat; fiecare tabel este tabel PowerPoint nativ. În runtime-ul offline din browser nu se poate importa/dezduplica literalmente un fișier PPTX. Pentru aceasta, exportul folosește geometria inspectată a exemplului furnizat `Prezentare in lucru.pptx`: fundalul alb cu blur BHB pentru conținut, fundaluri distincte pentru observații și profil comportamental, openerul cu gradient, fotografia alb-negru și simbolul/logo-ul BHB. Graficul de plajă folosește trei serii native (minim, mediană, maxim), echivalentul disponibil al graficului stock din șablon.

Nu sunt emise implicit zone/segmente, 9-box, personalitate, narațiune de client, recomandări, note de producție sau pagini „de completat”. Exporturile standard nu conțin datele necesare pentru aceste familii. Profilul comportamental este o pagină în două coloane, aliniată: top 3 după scor 2 și top 3 priorități după scor 0, fără coloane procentuale. Poate folosi opțional descriptorii Devplan; dacă fișierul Devplan este prezent, toate mapările observate trebuie să aibă textele pentru scorurile 0 și 2, altfel livrarea este blocată, fără fallback silențios.

## Ce recunoaște acum

- AC summary: antete `name`, `cod cp` și competențe cu scor 1–5.
- AC detailed: antete `name the person evaluated`, `cod ac`, `Competente` și `behavior`, cu scoruri 0–2. Media pe comportamente este convertită comparabil în 1–5 prin `1 + 2 × media(0–2)`.
- Devplan opțional: antete `competency`, `behavior`, `objective_text_score_0`, `objective_text_score_2` (sunt acceptate și aliasurile lizibile `Score 0` / `Score 2`); cheia este competență + comportament normalizată.
- 360: structură detectată și documentată ca viitor contract; nu se generează încă raport 360.

## Control, confidențialitate și limite

Validarea oprește livrabilul pentru antete necunoscute, nume/cod de evaluare lipsă, scor în afara scalei, identitate duplicată sau descriptor Devplan lipsă. Benchmarkul este editabil în interfață (implicit 2,75–3,50) și se aplică distribuției și benzii de pe graficul de plajă. Lipsa unui scor este avertisment: valoarea nu intră în medie, mediană, minim, maxim sau `n`; un participant fără niciun scor este exclus din toate calculele. Identitatea stabilă este **persoană + cod evaluare**, niciodată doar numele. Coloana sursă `CODE` este păstrată doar ca indicator de pregătire în audit și stare; nu se substituie cu `cod cp` sau `cod ac`.

Auditul XLSX păstrează amprenta de sursă, schema detectată, avertismente/blocaje, includerea sau excluderea, rândurile normalizate, calculele, agregatele de comportamente, distribuția și versiunea de calcul. Totul rulează în memoria browserului: fără upload, API, bază de date, analytics, browser storage, acces/modificare legacy, automatizare de click sau mesagerie. Nu păstrează fișierele și nu cere reintroducerea rosterului.

PDF-ul rămâne un pas local opțional din PowerPoint, nu o condiție de lucru. Pentru tabele sau grafice cu foarte multe competențe, operatorul poate ajusta dimensiunile și pozițiile direct în PPTX.

## Dezvoltare și recuperare

`npm run verify` reconstruiește `deploy/` și rulează contractele sintetice, inclusiv verificarea părților native de chart/workbook PptxGenJS. Nu introduce date client în teste sau în repository. Dacă o selecție este greșită, selectează din nou fișierul cu același nume pentru a-l înlocui sau folosește „Șterge sesiunea”; nu există date persistente de curățat.

Dependențele browser sunt livrate local în `src/assets/vendor/`: SheetJS, JSZip și PptxGenJS, cu licențele aferente. Nicio dependență nu este încărcată din rețea.
