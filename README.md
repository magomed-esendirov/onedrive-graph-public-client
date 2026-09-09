# OneDrive через Microsoft Graph: подключение публичного клиента

Как подключить свой OneDrive к собственному приложению или скрипту через Microsoft Graph API — **без `client_secret`**, и какие грабли мы собрали по дороге, пока это заработало.

## TL;DR — рабочая конфигурация

- Entra ID → App registrations → New registration
- **Supported account types:** `Accounts in any organizational directory (Any Microsoft Entra ID directory - Multitenant) and personal Microsoft accounts` — если входите личным аккаунтом (`@outlook.com`, `@hotmail.com`, `@live.com`)
- **Платформа:** `Mobile and desktop applications` (Native-клиент, public client — секрета нет и не нужен)
- **Redirect URI:** `https://login.microsoftonline.com/common/oauth2/nativeclient` (для device code flow URI не используется, но сама платформа должна быть добавлена)
- **Authentication → Allow public client flows: Yes**
- **API permissions (Delegated):** `User.Read`, `Files.ReadWrite`
- **Authority/endpoint:** `https://login.microsoftonline.com/common` (multitenant + personal) или `/consumers` (только personal) — **не** tenant-specific URL
- **Scopes:** `offline_access User.Read Files.ReadWrite`

## Пошаговая регистрация приложения

1. Открыть https://entra.microsoft.com → **Identity** → **Applications** → **App registrations** → **New registration**.
2. **Name** — любое, например `My OneDrive App`.
3. **Supported account types** — выбрать `Accounts in any organizational directory (Any Microsoft Entra ID directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)`.
4. Redirect URI на этом шаге можно пропустить — платформу добавим дальше.
5. **Register** → на странице Overview скопировать **Application (client) ID** — это ваш `client_id`.
6. Слева **Authentication** → **Add a platform** → **Mobile and desktop applications** → отметить `https://login.microsoftonline.com/common/oauth2/nativeclient` → **Configure**.
7. На той же странице внизу: **Allow public client flows → Yes** → **Save**.
8. Слева **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → отметить `User.Read` и `Files.ReadWrite` → **Add permissions**.
9. Для делегированных разрешений и личного аккаунта admin consent не требуется — можно сразу пользоваться.

## Почему не Web и не SPA: наши ошибки

Мы перебрали три типа платформы, прежде чем OAuth заработал. Каждая давала свою характерную ошибку — по ним легко понять, что не так.

### Ошибка 1: платформа Web → `AADSTS70002`

Если в **Authentication** добавлена платформа **Web**, приложение считается confidential client. При обмене authorization code на токен Microsoft требует секрет:

```
AADSTS70002: The request body must contain the following parameter:
'client_assertion' or 'client_secret'.
```

Для личного скрипта или агента хранить секрет не хочется и не нужно. Лечение: убрать платформу Web, сделать приложение публичным клиентом (секрет не требуется в принципе).

### Ошибка 2: платформа SPA → `AADSTS90023`

Убрали Web, поставили **Single-page application** — требование секрета пропало, но token endpoint стал отвечать:

```
AADSTS90023: Cross-origin token redemption is permitted only for the
'Single-Page Application' client-type.
```

Перевод: тип SPA разрешает обмен кода на токен **только из браузера** (запрос должен прийти с `Origin`-заголовком, cross-origin). Код, выполняющийся не в браузере (сервер, скрипт, агент), получает отказ. Настройками это не лечится — только сменой типа платформы.

### Решение: Native (Mobile and desktop applications)

Платформа **Mobile and desktop applications** — это public client без секрета, которому разрешён обмен кода вне браузера (authorization code + PKCE, device code flow). После перехода с SPA на Native OAuth прошёл с первого раза, больше ничего менять не пришлось.

### Ошибка 3: `AADSTS50020` — не тот tenant / не те типы аккаунтов

```
AADSTS50020: User account '...' from identity provider 'live.com' does not
exist in tenant '...'
```

Две возможные причины:

1. В регистрации стоит **Single tenant**, а вход выполняется личным аккаунтом. Лечение — выбрать **Multitenant + personal Microsoft accounts** (см. шаг 3 выше).
2. В коде используется tenant-specific endpoint `https://login.microsoftonline.com/<tenant-id>/...`. Для такого сценария нужен `https://login.microsoftonline.com/common` (multitenant + personal) или `https://login.microsoftonline.com/consumers` (только personal).

## Получение токена: device code flow

Самый простой способ для скрипта — не нужен ни веб-сервер, ни redirect URI:

```python
import msal

app = msal.PublicClientApplication(
    client_id="<ваш application (client) id>",
    authority="https://login.microsoftonline.com/common",
)
flow = app.initiate_device_flow(scopes=["User.Read", "Files.ReadWrite", "offline_access"])
print(flow["message"])  # открыть URL из сообщения, ввести код
result = app.acquire_token_by_device_flow(flow)
access_token = result["access_token"]
```

Токен живёт ~час; `offline_access` в скоупах даёт refresh token, `acquire_token_silent` продлевает сессию без участия пользователя.

## Основные вызовы Graph

База: `https://graph.microsoft.com/v1.0`, заголовок `Authorization: Bearer <access_token>`.

| Что | Запрос |
|---|---|
| Проверка токена | `GET /me` |
| Корень OneDrive | `GET /me/drive/root/children` |
| Поиск | `GET /me/drive/root/search(q='договор')` |
| Скачать файл | `GET /me/drive/root:/Documents/file.txt:/content` |
| Загрузить файл (≤ 4 МБ) | `PUT /me/drive/root:/Documents/file.txt:/content` |
| Удалить | `DELETE /me/drive/root:/Documents/file.txt` |

Пути вида `/me/drive/root:/Documents/Отчёт.pdf` нужно URL-кодировать (кроме `/`).

### Грабли со скачиванием: 302 на чужой хост

`GET ...:/content` **не возвращает файл сразу** — он отвечает **302** на хост вида `*.my.microsoftpersonalcontent.com`. Это pre-authenticated ссылка: токен уже вшит в URL.

**Не отправляйте туда заголовок `Authorization`.** Он там не нужен, а его присутствие ломает скачивание (у нас запрос просто висел до таймаута). Правильно:

1. Запросить `...:/content` с запретом авто-редиректов.
2. Взять `Location` из ответа.
3. Скачать `Location` **без** заголовка `Authorization`.

В `requests` это почти из коробки (библиотека сама снимает `Authorization` при редиректе на другой хост), но надёжнее обработать 302 вручную — см. функцию `download()` в `example/graph_client.py`.

### Нюанс поиска

`search(q='...')` ищет по поисковому индексу OneDrive, а не по живой файловой системе. Свежезагруженный файл появляется в выдаче **через несколько минут** — это нормально, не баг вашего кода.

## Пример

`example/graph_client.py` — минимальный рабочий клиент на `msal` + `requests`:

- вход через device code flow (с кэшем токена),
- `me`, список файлов в корне, поиск,
- загрузка, скачивание (с ручной обработкой 302), удаление.

```bash
pip install -r requirements.txt
export ONEDRIVE_CLIENT_ID="<ваш application (client) id>"
python example/graph_client.py
```

## Полезные ссылки

- Разбор `AADSTS90023` (cross-origin token redemption): https://learn.microsoft.com/en-us/answers/questions/588174/microsoft-devicecode-authentication-returning-inva
- Разбор `AADSTS50020` (аккаунт не найден в tenant): https://learn.microsoft.com/en-us/answers/questions/2203176/issue-with-microsoft-oauth-for-outlook-aadsts50020
