# Self-hosted Remote

## Phone collaboration in BigDogOffice

Open **Settings → Phone collaboration**. For Wi-Fi or cellular pairing, connect the desktop to a Remote service at an HTTPS address reachable from both devices. Complete desktop authorization before generating a pairing invitation.

Save **Start Remote with BigDog** to connect on the next launch. Automatic path selection tries direct LAN or P2P, TURN when available, and encrypted WebSocket relay. Choose **Server relay** to use the relay directly. The desktop shows the route selected by the connected phone.

The built-in development service listens on loopback at port 48300. Use an emulator or a device with an ADB reverse port forward for this mode.

Remote is a single-account coordination service for your own devices. It supplies authentication, signaling, relay, short-lived TURN credentials, encrypted offline delivery and execution ownership. Desktop model credentials and filesystem authority remain on the host.

## Local start

Use Node.js 24 and Docker for the container deployment:

```sh
npm ci
npm run remote:package
cp apps/remote-server/.env.example apps/remote-server/.env
```

Set your own `REMOTE_ACCOUNT` and a long random `REMOTE_PASSWORD` in that ignored file, then run:

```sh
docker compose --env-file apps/remote-server/.env -f apps/remote-server/compose.yaml up -d --build
```

The local service binds to `127.0.0.1:48300`. Do not expose the unencrypted development endpoint publicly. Keep the persistent volume; normal shutdown should not remove it.

To run without Docker, set `DSH_SERVER_ACCOUNT`, `DSH_SERVER_PASSWORD`, `DSH_SERVER_PUBLIC_URL` and `DSH_SERVER_DATA_FILE` privately, then run `npm run start -w @bigdogoffice/remote-server`.

## Desktop connection

Create a local JSON outside source control:

```json
{
  "serverUrl": "http://127.0.0.1:48300",
  "account": "your-account@example.com",
  "password": "YOUR_OWN_PASSWORD",
  "name": "My desktop",
  "pairingCode": "YOUR_RANDOM_INITIAL_INVITATION"
}
```

Set `BIGDOG_REMOTE_CONFIG` to that file before launching the desktop. Routine pairing then uses the desktop's mobile-collaboration settings. QR invitations do not contain model keys or the server password.

## Network and consistency

LAN/P2P → TURN → Relay is an adaptive encrypted transport policy. Direct setup still requires a reachable coordinator; an internet-independent LAN needs its own LAN coordinator. Authentication failure does not downgrade to plaintext.

For public deployment, supply your own domain, TLS configuration and TURN public address/secret using the example variables. `compose.yaml` provides a public profile using Caddy and coturn. Account/device state and encrypted courier SQLite data must be backed up consistently.

Stable device identity and request receipts survive ordinary restarts/re-pairing; uninstalling or clearing data loses device identity. Offline requests are accepted into an encrypted inbox, then revalidated by the host before execution. Per-device receipts are authoritative. Idle Session handoff uses an ownership epoch; do not synchronize live SQLite files as a multi-host execution protocol.
