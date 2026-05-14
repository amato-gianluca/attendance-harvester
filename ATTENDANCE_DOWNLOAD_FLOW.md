# Flusso Dettagliato di Download degli Attendance Report

Questo documento descrive la procedura utilizzata da `main.py` per scaricare gli attendance report da Microsoft Teams, con i dettagli delle singole API Graph chiamate.

## Panoramica del Flusso

Il processo di download segue questi step principali:
1. **Autenticazione** → Acquire token Microsoft Graph
2. **Inizializzazione client** → Creazione GraphClient
3. **Scoperta Teams** → Get all subscribed teams
4. **Filtraggio Teams** → Applica regex sul nome team
5. **Scoperta Canali** → Get all team channels
6. **Scoperta Meeting** → Query calendar per meetings in date range
7. **Risoluzione Meeting** → Associa calendar event con online meeting
8. **Estrazione Attendance** → Download attendance reports e records
9. **Esportazione** → Salva in CSV/JSON

---

## Step 1: Autenticazione

**Funzione:** `acquire_access_token_from_config()` in `main.py`

### Flow di autenticazione:

1. **Inizializza Authenticator** (da `src/auth.py`)
   - Parametri: client_id, authority, scopes, auth_mode ("public" o "confidential")
   - Se "public": device code flow (browser-based)
   - Se "confidential": client credentials flow (app-only token)

2. **Acquisisce token**
   ```python
   authenticator.acquire_token()
   ```
   - Restituisce: Bearer token per autorizzare tutte le richieste Graph seguenti
   - Token salvato in cache locale per riutilizzo

---

## Step 2: Inizializzazione GraphClient

**Funzione:** `run_harvest()` in `main.py`

### Creazione client:

```python
graph_client = GraphClient(
    access_token=access_token,
    max_retries=3,
    retry_backoff_factor=2,
    timeout=30,
    user_id=target_user_id,  # Solo se mode=confidential
    metadata_cache_file="cache/teams_channels.json"
)
```

- **max_retries**: Numero di tentativi per errori transitori (rate limit, 5xx)
- **retry_backoff_factor**: Attesa esponenziale (2^attempt secondi)
- **timeout**: 30 secondi per singola richiesta
- **user_id**: Richiesto in confidential mode per scope le API a uno specifico utente
- **metadata_cache_file**: Cache locale di team/channels per ridurre API calls

---

## Step 3: Scoperta Teams

**API: GET /me/joinedTeams** (o **/users/{id}/joinedTeams** in confidential mode)

**Funzione:** `GraphClient.get_joined_teams()`

### Procedura:

1. **Primo endpoint della paginazione:**
   ```
   GET https://graph.microsoft.com/v1.0/me/joinedTeams
   ```

2. **Gestione paginazione:**
   - Response contiene array `value` con team objects
   - Se presente `@odata.nextLink`: fare richiesta a quel URL per pagina successiva
   - Continua finché non ci sono più `@odata.nextLink`

3. **Response payload per ogni team:**
   ```json
   {
     "id": "team-guid",
     "displayName": "Team Name",
     "description": "Team description"
   }
   ```

4. **Opzionale - Associated Teams** (team condivisi, shared channels)
   ```
   GET /me/teamwork/associatedTeams
   ```
   - Restituisce array di associated team info
   - Aggiunge team non ancora presenti in joinedTeams

---

## Step 4: Filtraggio Teams

**Funzione:** `TeamFilter.filter_teams()` in `src/team_filter.py`

### Procedura:

1. Applica regex pattern configurato (es: `^Didattica.*`) al campo `displayName`
2. Mantiene solo i team che matchano
3. Valida che `id` sia un GUID valido

**Output:** Lista filtrata di team objects

---

## Step 5: Scoperta Canali per Team

**API: GET /teams/{teamId}/channels** o **/teams/{teamId}/primaryChannel**

**Funzione:** `GraphClient.get_team_channels()` o `GraphClient.get_team_primary_channel()`

### Configurazione:

Se config.meetings.general_channel_only = true:
- Chiama: `GET /teams/{teamId}/primaryChannel`
- Restituisce il canale "Generale" (General)

Se false:
- Chiama: `GET /teams/{teamId}/channels`
- Restituisce TUTTI i canali del team (con paginazione)

### Caching:

- I dati dei canali sono cachati in memoria e in `cache/teams_channels.json`
- La cache è sincronizzata con i team filtrati (rimuove entry di team non più matchati)

### Response payload per channel:

```json
{
  "id": "channel-guid",
  "displayName": "Channel Name",
  "description": "Channel description"
}
```

### Output di Step 3-5:

```python
teams_with_channels = [
    {
        "team": {"id": "team1", "displayName": "Team A"},
        "channel": {"id": "ch1", "displayName": "General"}
    },
    {
        "team": {"id": "team1", "displayName": "Team A"},
        "channel": {"id": "ch2", "displayName": "Announcements"}
    }
]
```

---

## Step 6: Scoperta Meeting nel Date Range

**API: GET /me/calendarView** (scoped per date range)

**Funzione:** `MeetingResolver.get_meetings_in_date_range()` in `src/meeting_resolver.py`

### Procedura:

1. **Calcola date range:**
   - Start: `now - lookback_days`
   - End: `now + lookahead_days`
   - Formattate in ISO 8601 (es: "2026-05-14T10:00:00+00:00")

2. **API Call:**
   ```
   GET /me/calendarView?
       startDateTime=2026-05-01T00:00:00Z&
       endDateTime=2026-05-14T23:59:59Z&
       $select=id,subject,start,end,isOnlineMeeting,onlineMeetingProvider,onlineMeeting,organizer,location
   ```

3. **Filtro a Teams meetings:**
   - Mantiene solo event con:
     - `isOnlineMeeting = true`
     - `onlineMeetingProvider = "teamsForBusiness"`

4. **Gestione paginazione:**
   - Se response ha `@odata.nextLink`, fetcha pagina successiva
   - Accumula tutti gli event in un'unica lista

### Response payload per meeting event:

```json
{
  "id": "event-id",
  "subject": "Meeting Title",
  "start": {"dateTime": "2026-05-14T10:00:00", "timeZone": "UTC"},
  "end": {"dateTime": "2026-05-14T11:00:00", "timeZone": "UTC"},
  "isOnlineMeeting": true,
  "onlineMeetingProvider": "teamsForBusiness",
  "onlineMeeting": {
    "joinUrl": "https://teams.microsoft.com/l/meetup-join/19:meeting@thread.tacv2/..."
  },
  "organizer": {"emailAddress": {"address": "organizer@example.com"}}
}
```

---

## Step 7: Associazione Calendar Event con Team/Channel

**Funzione:** `MeetingResolver._match_event_contexts_from_join_url()` in `src/meeting_resolver.py`

### Procedura:

1. **Estrae thread ID dal join URL:**
   - Parse URL path per cercare pattern `*@thread.tacv2` oppure `*@thread.v2`
   - Se non trovato, estrae da query parameter `context` (decodifica JSON)

2. **Matchizza con team/channel:**
   - Per ogni team/channel combination, estrae `channel.id`
   - Se `channel.id == url_thread_id`: match trovato
   - Raccoglie tutti i match

3. **De-duplicazione:**
   - Rimuove duplicati per coppia team+channel

### Output:

```python
matched_contexts = [
    {
        "team": {"id": "team-guid", "displayName": "Didattica 2025/26"},
        "channel": {"id": "channel-guid", "displayName": "General"}
    }
]
```

---

## Step 8: Risoluzione Online Meeting ID

**API: GET /me/onlineMeetings (con filter) oppure GET /me/calendarView**

**Funzione:** `MeetingResolver.resolve_online_meeting()` in `src/meeting_resolver.py`

### Procedura:

1. **Try #1 - Fetch online meeting object dal join URL (per meeting organizzati):**
   ```
   GET /me/onlineMeetings?$filter=joinWebUrl eq '{join_url}'
   ```
   - Se trovato: restituisce l'online meeting object con `id` (meeting ID)

2. **Try #2 - Se #1 fallisce, prova con owner fallback:**
   - Per ogni team matched, ottiene i team owners:
     ```
     GET /groups/{teamId}/owners
     ```
   - Per ogni owner, riprova la risoluzione online meeting con:
     ```
     GET /users/{ownerId}/onlineMeetings?$filter=joinWebUrl eq '{join_url}'
     ```

3. **Try #3 - Fallback per meeting non organizzati:**
   - Se entrambi i tentativi precedenti falliscono (normale per meeting a cui partecipi solo)
   - Crea un minimal meeting object contenente il join URL
   - Questo consente comunque di tentare il fetch degli attendance report (fallirà con 404 se non organizzato)

### Response payload online meeting (Try #1 o #2):

```json
{
  "id": "meeting-guid",
  "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/...",
  "chatInfo": {
    "threadId": "channel-id@thread.tacv2"
  }
}
```

### Output:

```python
online_meeting = {
    "id": "meeting-guid",
    "joinWebUrl": "...",
    "_event": {
        "subject": "Meeting Title",
        "start": {...},
        "end": {...},
        "organizer": {...}
    },
    "_resolved_user_id": "user-guid"  # Se resolved con owner fallback
}
```

---

## Step 9: Estrazione Attendance Report per Meeting

**API: GET /me/onlineMeetings/{meetingId}/attendanceReports**

**Funzione:** `MeetingResolver.get_channel_attendance()` → `GraphClient.get_attendance_reports()`

### Procedura per ogni online meeting:

1. **Fetch attendance reports:**
   ```
   GET /me/onlineMeetings/{meetingId}/attendanceReports
   ```
   - Se era risotto con owner fallback user: esegui con `/users/{ownerId}/...`
   - Se response è 404: il meeting non è stato organizzato da questo utente (normale)

2. **Gestione paginazione:**
   - Se presence di `@odata.nextLink`: accumula tutte le pagine

3. **Deduplicazione report:**
   - Per ogni report trovato, salva `report.id`
   - Se stesso report viene restituito da multiple API calls: mantieni una sola copia

### Response payload per report:

```json
{
  "id": "report-guid",
  "meetingStartDateTime": "2026-05-14T10:00:00Z",
  "meetingEndDateTime": "2026-05-14T11:00:00Z",
  "totalParticipantCount": 25
}
```

---

## Step 10: Estrazione Attendance Record per Report

**API: GET /me/onlineMeetings/{meetingId}/attendanceReports/{reportId}/attendanceRecords**

**Funzione:** `MeetingResolver.get_channel_attendance()` → `GraphClient.get_attendance_records()`

### Procedura per ogni report:

1. **Fetch attendance records:**
   ```
   GET /me/onlineMeetings/{meetingId}/attendanceReports/{reportId}/attendanceRecords
   ```
   - Usa lo stesso user_id che ha restituito il report (può essere owner fallback)

2. **Gestione paginazione:**
   - Accumula tutte le pagine di record

### Response payload per attendance record:

```json
{
  "id": "record-guid",
  "emailAddress": "student@example.com",
  "displayName": "Student Name",
  "role": "participant",
  "totalAttendanceInSeconds": 1800,
  "identity": {
    "id": "user-guid"
  },
  "attendanceIntervals": [
    {
      "joinDateTime": "2026-05-14T10:00:00Z",
      "leaveDateTime": "2026-05-14T10:30:00Z"
    }
  ]
}
```

---

## Step 11: Mappatura Report → Meeting

**Funzione:** `MeetingResolver._select_best_meeting_for_report()` in `src/meeting_resolver.py`

### Procedura:

1. **Filtra candidate matching:**
   - Mantieni solo meeting che hanno effettivamente restituito questo report
   - (Un report può essere mappato solo a meeting che lo hanno generato)

2. **Time-based matching:**
   - Calcola distanza temporale tra:
     - `report.meetingStartDateTime` ↔ `meeting._event.start`
     - `report.meetingEndDateTime` ↔ `meeting._event.end`
   - Seleziona il meeting con distanza minima

3. **Fallback:**
   - Se nessun time match possibile: usa il primo candidate (generalmente esiste un solo candidate)

### Output:

```python
selected_meeting = {
    "meeting_id": "meeting-guid",
    "meeting_info": {
        "subject": "Meeting Title",
        "start": {...},
        "end": {...}
    },
    "teams_context": [
        {
            "team": {...},
            "channel": {...}
        }
    ]
}
```

---

## Step 12: Assemblaggio Payload di Attendance

**Funzione:** `MeetingResolver.get_channel_attendance()` in `src/meeting_resolver.py`

### Payload finale per ogni report:

```python
attendance_data_item = {
    "meeting_id": "meeting-guid",
    "meeting_info": {
        "subject": "Meeting Title",
        "start": {...},
        "end": {...},
        "organizer": {...}
    },
    "report_id": "report-guid",
    "report_data": {
        "id": "report-guid",
        "meetingStartDateTime": "...",
        "meetingEndDateTime": "...",
        "totalParticipantCount": 25
    },
    "attendance_records": [
        {
            "emailAddress": "student1@example.com",
            "displayName": "Student 1",
            "totalAttendanceInSeconds": 1800,
            "attendanceIntervals": [...]
        },
        {
            "emailAddress": "student2@example.com",
            "displayName": "Student 2",
            "totalAttendanceInSeconds": 2700,
            "attendanceIntervals": [...]
        }
    ],
    "teams_context": [
        {
            "team": {"id": "team-guid", "displayName": "Team Name"},
            "channel": {"id": "channel-guid", "displayName": "Channel Name"}
        }
    ],
    "source_meeting_id": "meeting-guid"
}
```

---

## Step 13: Salvataggio JSON Export

**Funzione:** `AttendanceExporter.export_batch()` in `src/exporter.py`

### Procedura:

1. **Per ogni attendance_data_item:**
   - Costruisce filename usando pattern: `{team_name}_{channel_name}_{meeting_date}_{meeting_id}_{report_id}_attendance.json`
   - Salva full payload in `output/json/{team_name}/filename.json`

2. **Directory structure:**
   ```
   output/json/
   ├── TEAM_NAME_A [a] [18]/
   │   ├── Team_A_General_2026-05-14_meeting1_report1_attendance.json
   │   └── Team_A_General_2026-05-14_meeting2_report2_attendance.json
   └── TEAM_NAME_B [b] [12]/
       └── Team_B_Announcements_2026-05-13_meeting3_report3_attendance.json
   ```

---

## Step 14: Generazione CSV Export

**Funzione:** `AttendanceExporter.export_batch()` in `src/exporter.py`

### Procedura per ogni attendance_data_item:

1. **Valida durata minima:**
   - Se `config.output.min_csv_report_duration_seconds > 0`:
     - Calcola durata report: `meetingEndDateTime - meetingStartDateTime`
     - Se durata < min: **salta** export CSV (ma JSON già salvato)

2. **Costruisce dataset CSV:**
   - **Sezione 1 - Meeting Info:**
     ```
     Team: TEAM_NAME
     Channel: CHANNEL_NAME
     Subject: Meeting Title
     Date: 2026-05-14
     Start Time: 10:00 AM
     End Time: 11:00 AM
     Total Participants: 25
     ```

   - **Sezione 2 - Participants with attendance details:**
     ```
     Name, First Join, Last Leave, In-Meeting Duration, Email, Participant ID, Role
     Student 1, 10:00:00, 10:30:00, 00:30:00, student1@example.com, id1, participant
     Student 2, 10:05:00, 11:00:00, 00:55:00, student2@example.com, id2, participant
     ```

3. **Salva file:**
   - Naming identico a JSON
   - Path: `output/csv/{team_name}/filename.csv`

---

## Step 15: Upload CSV a SharePoint (Opzionale)

**Funzione:** `SharePointCSVUploader.upload_files()` in `src/sharepoint_csv_uploader.py`

### Procedura (se configurato):

1. **Resolve SharePoint site:**
   ```
   GET /sites/{site_hostname}:/sites/{site_path}
   ```

2. **Resolve root drive:**
   ```
   GET /sites/{site_id}/drives
   GET /drives/{drive_id}/root
   ```

3. **Per ogni CSV file:**
   - Crea folder structure in SharePoint: `/folder_path/{team_name}/`
   - Upload file:
     ```
     PUT /drives/{drive_id}/root:/path/file.csv:/content
     ```
   - Rename folder adding suffix ` [open]`

---

## Riassunto API Calls Totali

| Funzione | API Endpoint | Quando |
|----------|-----------|--------|
| Get joined teams | GET /me/joinedTeams | Sempre |
| Get associated teams | GET /me/teamwork/associatedTeams | Se include_associated_teams=true |
| Get team owners | GET /groups/{teamId}/owners | Per meeting fallback resolution |
| Get team channels | GET /teams/{teamId}/channels | Se general_channel_only=false |
| Get primary channel | GET /teams/{teamId}/primaryChannel | Se general_channel_only=true |
| Get calendar view | GET /me/calendarView | Sempre (con date range) |
| Get online meeting | GET /me/onlineMeetings (filter) | Per ogni calendar event |
| Get online meeting (owner) | GET /users/{ownerId}/onlineMeetings (filter) | Se meeting fallback per owner |
| Get attendance reports | GET /me/onlineMeetings/{id}/attendanceReports | Per ogni online meeting risotto |
| Get attendance records | GET /me/onlineMeetings/{id}/attendanceReports/{id}/attendanceRecords | Per ogni attendance report |
| SharePoint site | GET /sites/{id} | Se SharePoint upload abilitato |
| SharePoint drive | GET /drives/{id}/root | Se SharePoint upload abilitato |
| Upload CSV | PUT /drives/{id}/root:/path:/content | Se SharePoint upload abilitato |

---

## Error Handling e Retry Logic

### Retry Strategy (in `GraphClient._make_request()`):

1. **Rate limit (429):**
   - Aspetta `Retry-After` header, default 2^attempt secondi
   - Riprova fino a max_retries

2. **Server error (5xx):**
   - Backoff esponenziale: 2^attempt secondi
   - Riprova fino a max_retries

3. **Not found (404):**
   - **Non** considerato errore critico
   - Ritorna 404 response (usato per non-organized meetings)

4. **Client error (4xx) escluso 404:**
   - Raise GraphAPIError immediato
   - **Nessun retry**

5. **Network error (timeout, connection):**
   - Backoff esponenziale
   - Riprova fino a max_retries

---

## Performance Notes

- **Caching metadata:** Team/channel metadata cachato per ridurre API calls
- **Processed report tracking:** Report già esportati vengono skippati se `--skip-processed=True`
- **Paginazione:** Tutte le API implementano paginazione via `@odata.nextLink`
- **Metrica:** Per 1 team con 10 meeting organizzati, ~30-40 API calls totali
