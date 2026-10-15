# OneDrive-klantenportaal

Kleine Next.js App Router-app met TypeScript en Tailwind. Eén directe submap van één privé-OneDrive-hoofdmap is één project. Geen database, synchronisatiejobs of klantaccounts. Graph-integratie staat in `lib/graph.ts`; de UI in `app/[projectId]/page.tsx`. Er is geen publieke projectenlijst. Nieuwe bestanden en mappen worden bij de volgende aanvraag rechtstreeks uit Graph gelezen, inclusief paginering.

## Lokaal starten

Gebruik Node.js 22.18+ (ook getest met Node 24).

```sh
npm ci
cp .env.example .env.local
# Vul .env.local lokaal in; nooit committen.
npm run dev
```

Configuratie:

| Variabele | Betekenis |
| --- | --- |
| `AZURE_TENANT_ID` | Directory-ID van de Microsoft 365-tenant |
| `AZURE_CLIENT_ID` | Application/client-ID van de Entra-app |
| `AZURE_CLIENT_SECRET` | Secretwaarde van die app; niet het secret-ID |
| `ONEDRIVE_FOLDER_URL` | Optioneel: volledig zakelijk browseradres van de hoofdmap; zie automatische ontdekking hieronder |
| `ONEDRIVE_DRIVE_ID` | ID van de zakelijke OneDrive-drive |
| `ONEDRIVE_ROOT_FOLDER_ID` | Item-ID van de hoofdmap, bijvoorbeeld Klantenportaal |
| `PROJECT_LINK_SECRET` | Vaste geheime willekeurige sleutel van minstens 32 bytes; genereer met `openssl rand -hex 32` |
| `APP_BASE_URL` | Publieke oorsprong, bijvoorbeeld `https://app.domein.nl`; voor CLI-links |

Alle variabelen blijven server-side. Gebruik nooit het prefix `NEXT_PUBLIC_` voor deze waarden.

## Hoofdmap instellen met het browseradres

Als alternatief voor de twee ID's kunt u in Vercel **`ONEDRIVE_FOLDER_URL`** invullen met het volledige adres van de hoofdmap uit de OneDrive-browser. Bijvoorbeeld (fictief):

```text
https://example-my.sharepoint.com/my?id=%2Fpersonal%2Farchitect_example_nl%2FDocuments%2FApp
```

Laat bij deze optie `ONEDRIVE_DRIVE_ID` en `ONEDRIVE_ROOT_FOLDER_ID` leeg of verwijder ze. Twee ingevulde ID's hebben voorrang; één losse ID is een configuratiefout. Redeploy na het wijzigen van Vercel-variabelen. `APP_BASE_URL` is het publieke appdomein, niet het OneDrive-adres.

De server bepaalt via Graph de site, de standaarddocumentbibliotheek en de hoofdmap. Er zijn geen interactieve gebruikersaanmeldingen of `/me`-aanvragen. Zowel de OneDrive-weergave met `?id=` als directe mapadressen onder `/personal/`, `/sites/` en `/teams/` worden ondersteund; alleen HTTPS op `*.sharepoint.com`. Anonieme deel-URL's, persoonlijke OneDrive en andere documentbibliotheken worden niet ondersteund. Het ingestelde browseradres wordt nooit rechtstreeks opgehaald. Alleen de gevonden ID's worden maximaal 60 seconden in het proces onthouden; projectinhoud wordt steeds opnieuw gelezen. Houd het hoofdmapadres actueel als de hoofdmap zelf wordt hernoemd. Projectlinks blijven gelijk zolang drive-ID, projectmap-ID en de serversleutel gelijk blijven.

**Rechten blijven een vereiste.** Automatische ontdekking vraagt site- en bibliotheekmetadata op. Alleen `Files.SelectedOperations.Selected` met een maptoekenning is daarvoor niet voldoende. Voor deze optie moet een beheerder expliciet **`Sites.Selected` application permission**, admin consent én een **`read`-toekenning op de betreffende site** instellen. Bij OneDrive gaat het om de persoonlijke zakelijke SharePoint-site van de eigenaar. Dat geeft de runtime-app ruimere leesrechten op die geselecteerde site dan alleen de map. Beoordeel die verruiming bewust; de portal zelf begrenst aanvragen nog steeds tot de ingestelde hoofdmap en het gekozen project. Zonder site-toekenning kan automatische ontdekking niet werken. De app vraagt zelf geen extra rechten aan, maakt niets publiek en voert geen schrijfbewerkingen uit.

Voor de striktste toegang tot alleen de hoofdmap blijft de configuratie met beide ID's en `Files.SelectedOperations.Selected` hieronder de aanbevolen optie. Een verkeerde tenant, ontbrekende admin consent of ontbrekende read-toekenning wordt niet opgelost door alleen een URL in te vullen. Site-toekenningen worden beheerd via een afzonderlijke beheercontext, bijvoorbeeld `POST /sites/{site-id}/permissions` met `roles: ["read"]` en `grantedToIdentities: [{"application": {"id": "{client-id}"}}]`; beheerrechten horen nooit in de portal-app.

## Microsoft Entra en Microsoft Graph

Deze implementatie gebruikt **OneDrive for Business binnen Microsoft 365**, met app-only client credentials. Een persoonlijke OneDrive/Microsoft-account ondersteunt deze aanpak niet: daarvoor is een andere implementatie met delegated OAuth, toestemming van de eigenaar en veilige opslag/vernieuwing van refresh tokens nodig.

1. Maak in Microsoft Entra ID een single-tenant appregistratie. Een redirect-URI is voor client credentials niet nodig. Noteer tenant-ID en client-ID.
2. Maak onder Certificates & secrets een client secret; bewaar de waarde veilig en plan rotatie vóór de vervaldatum.
3. Voeg Microsoft Graph **Application permission `Files.SelectedOperations.Selected`** toe en laat een tenantbeheerder admin consent geven. Dit geeft op zichzelf nog geen toegang tot bestanden. Voeg géén `Files.Read.All`, `Files.ReadWrite.All` of andere tenantbrede rechten toe aan de runtime-app.
4. Zoek met een **afzonderlijke beheercontext** de drive-ID (`GET /users/{user-id}/drive`) en het item-ID van Klantenportaal (`GET /drives/{drive-id}/root:/Klantenportaal`). Bij deze variant doet de portal-app geen ontdekking: met alleen mapgebonden rechten kan deze niet zomaar alle gebruikers/drives ontdekken. Gebruik `/me` niet met app-only authenticatie.
5. Ken die runtime-app een expliciete **read**-toekenning op de hoofdmap toe. Een beheerder doet dit eenmalig via Graph, met een apart bevoegd beheeraccount of provisioning-app. De beheercontext moet de door Microsoft vereiste rechten voor het beheren van deze toekenning hebben (bijvoorbeeld `Sites.FullControl.All` application met admin consent); bewaar deze beheercredentials nooit in de portal. Voorbeeld:

```http
POST https://graph.microsoft.com/v1.0/drives/{drive-id}/items/{hoofdmap-item-id}/permissions
Content-Type: application/json
Authorization: Bearer {beheer-token}

{
  "roles": ["read"],
  "grantedToV2": {
    "application": { "id": "{client-id-van-de-portal-app}" }
  }
}
```

De toekenning geldt voor de hoofdmap en overervende onderliggende inhoud, niet voor de rest van OneDrive. Controleer dit daadwerkelijk met het runtime-token. Afwijkende overerving op bestaande of verplaatste submappen kan toegang blokkeren; herstel dat in Microsoft 365. De geselecteerde maptoekenning kan de overerving op die map verbreken; laat de beheerder de gevolgen controleren. Maak geen anonieme OneDrive-deellinks.

Controleer met het runtime-token `GET /drives/{drive-id}/items/{hoofdmap-item-id}/children` en een bestandsdownload daarbinnen. Controleer ook dat een bestand buiten de hoofdmap wordt geweigerd. Wanneer geselecteerde maprechten door tenantbeleid niet beschikbaar zijn, laat de beheerder dit oplossen; verruim de runtime-app niet stilzwijgend naar tenantbrede toegang. Een geselecteerde site met read-toekenning (`Sites.Selected`) is eventueel een bewust te beoordelen ruimer alternatief.

Officiële documentatie: [Selected permissions](https://learn.microsoft.com/en-us/graph/permissions-selected-overview), [driveItem-toekenning](https://learn.microsoft.com/en-us/graph/api/driveitem-post-permissions?view=graph-rest-1.0), [client credentials](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow).

## Projectlinks kopiëren

```sh
npm run projects
```

Dit lokale CLI-commando leest `.env.local` en toont per directe submap de naam en URL. Het is geen webendpoint. De ID is HMAC-SHA256 van drive-ID en map-ID, met de geheime serversleutel (43 base64url-tekens). Een hernoeming verandert de link niet. Nieuwe projectmappen verschijnen vanzelf. Verplaatsing naar een andere drive of opnieuw aanmaken kan het map-ID en dus de link veranderen.

**Iedereen met de projectlink kan dat project bekijken en downloaden.** Behandel URL's als toegangssleutels, ook in browsergeschiedenis, hostinglogs, supportberichten en analytics. Voeg geen externe analytics of externe pagina-assets toe die links kunnen vastleggen. Er zijn geen individuele intrekkingen zonder database: verwijder/verplaats de projectmap buiten de hoofdmap om de toegang te stoppen, of roteer `PROJECT_LINK_SECRET` om alle links te vervangen. Reeds gedownloade documenten kunnen niet worden teruggehaald.

## Beveiliging en werking

- Elke aanvraag valideert de link opnieuw tegen de huidige directe projectmappen. Elke aangevraagde map, preview en download wordt via de echte parent-keten aan dat project gekoppeld. Andere projecten, externe snelkoppelingen en remote items worden geweigerd.
- Bestanden worden via een serverroute gestreamd. Graph access tokens en vooraf geauthenticeerde download-URL's worden nooit naar de browser gestuurd. De server volgt uitsluitend de door Graph verstrekte HTTPS-downloadlocatie, zonder Graph Authorization-header. Byte-ranges worden doorgegeven voor PDF's.
- Alleen PDF en veilige rasterafbeeldingsformaten krijgen een inline preview. SVG, HTML en andere bestanden zijn uitsluitend attachments met `nosniff`. PDF-weergave hangt af van browserondersteuning; downloaden blijft beschikbaar.
- Pagina's en bestanden gebruiken geen gedeelde inhoudscache; downloadresponses zijn `private, no-store`. Referrers en indexering zijn uitgeschakeld. OneDrive blijft privé, de app exposeert uitsluitend GET/HEAD voor lezen.
- Manifest, 192/512px-iconen en Apple-icoon zijn aanwezig. Installeren via de browser kan via HTTPS. Geen serviceworker of offline documentcache: privébestanden worden niet bewust offline bewaard. Iedere projectpagina verwijst naar een eigen manifest met de projectlink als `id` en `start_url`. Het homescreen-icoon opent daarmee altijd die projectmap, ook wanneer u het toevoegt vanuit een submap of preview. Open op iOS de projectlink in Safari en kies Deel → Zet op beginscherm. Een eerder toegevoegd icoon moet u verwijderen en opnieuw toevoegen om het nieuwe manifest te gebruiken. De neutrale startpagina heeft een eigen algemeen manifest.
- Geen atomic snapshot: OneDrive-wijzigingen tijdens een aanvraag kunnen tijdelijk een ontbrekend item of foutmelding opleveren. Bij veel projecten kost het zoeken op HMAC meerdere Graph-pagina's. Dit is bewust eenvoudig gehouden.

## Controles

```sh
npm test             # Graph-mocks: projectgrenzen, HMAC, paginering en downloads
npm run lint
npm run typecheck
npm run build
npm run test:http    # productiebuild vereist; echte Next-server met gesimuleerd Graph
npm start
```

Mocks vervangen geen tenanttest. Controleer na het invullen van echte credentials: CLI, twee verschillende projecten, geneste mappen, PDF/afbeelding, downloads, hernoemen, nieuw project en geweigerde toegang buiten de map. Een foutmelding bij ontbrekende configuratie of Graph-storing toont geen Microsoft-response of credentials.

## Vercel

1. Importeer deze repository als Next.js-project in Vercel. Gebruik Node 24.x, standaard `npm ci` en `npm run build`.
2. Voeg de variabelen uit `.env.example` toe onder Environment Variables; voer secrets uitsluitend daar in. Gebruik aparte secrets en liefst aparte OneDrive-mappen/apps voor preview en productie. Zet `APP_BASE_URL` op het definitieve HTTPS-domein en houd `PROJECT_LINK_SECRET` stabiel tussen deployments.
3. Deploy. Er zijn geen build-time Graph-credentials nodig voor prerendering; projectroutes worden op aanvraag uitgevoerd. Controleer de echte tenantflow na deployment.
4. Voer de CLI lokaal uit met dezelfde drive, hoofdmap en projectlinksleutel als productie en de productie-`APP_BASE_URL`.

De server moet `login.microsoftonline.com`, `graph.microsoft.com` en de door Microsoft uitgegeven OneDrive/SharePoint-downloadhost van de tenant kunnen bereiken. Er zijn geen echte credentials of klantdocumenten meegeleverd. Let bij grote bestanden op de actuele Vercel Functions-limieten voor streaming, duur en verkeer. Test representatieve grote PDF's op het gekozen abonnement; deze MVP omzeilt hostinglimieten niet met publieke downloadlinks.

## Deelbare demonstratie

Open `/zhq17sD-hJ324XCT1yKvb2IEsErxs_4c` op het gedeployde domein. Deze vaste voorbeeldlink werkt zonder Microsoft-configuratie en toont uitsluitend fictieve PDF's en een schematische afbeelding. Mappen, previews en downloads werken met dezelfde UI als echte projecten. Er is een duidelijke demonstratiemelding. De bestanden in `public/demo` zijn openbare voorbeelden en bevatten geen klantgegevens. De demo verleent geen toegang tot OneDrive en verschijnt niet in de CLI-projectenlijst. Verwijder de demo-integratie en `public/demo` wanneer deze niet meer gewenst is.
