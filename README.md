# Discord House → Excel Bot

## 1. Install Python
Use Python 3.11 or newer.

## 2. Install dependencies

Windows:
```bat
py -3 -m pip install -r requirements.txt
```

## 3. Configure the bot

Copy `.env.example` to `.env` and put your Discord bot token in:

```env
DISCORD_TOKEN=YOUR_BOT_TOKEN
HOUSE_CHANNEL_ID=0
```

`HOUSE_CHANNEL_ID=0` means the bot can automatically process matching house-list messages in any channel it can read.

If you want only one channel, enable Developer Mode in Discord, right-click the channel, choose Copy Channel ID, and put that number in `HOUSE_CHANNEL_ID`.

## 4. Put your Excel template in place

The supplied template is already included as:

```text
templates/template.xlsx
```

You can replace it with your own template, or upload a new one with Discord:

```text
/settemplate
```

Attach your `.xlsx` file.

The workbook must contain a sheet named `Sheet1`.

## 5. Create the Discord bot

Go to the Discord Developer Portal:
https://discord.com/developers/applications

Create an application, then create a Bot.

Turn on:
- Message Content Intent

Invite the bot using OAuth2 → URL Generator with:
- bot
- applications.commands

Recommended permissions:
- View Channels
- Send Messages
- Attach Files
- Read Message History

## 6. Start the bot

Windows:
```bat
py -3 bot.py
```

You should see:

```text
Logged in as YourBot
Synced 3 slash command(s).
```

## 7. Use it

Paste your house list into the configured Discord channel.

Example:

```text
✨🌷 𝟱𝟱 𝗣𝗘𝗦𝗢𝗦 𝗛𝗢𝗨𝗦𝗘𝗦 🌷✨

189. Fluff Station V1 — ₱55
52k queenslander
cutecoress file name

190. Fluff Station V2 — ₱55
53k queenslander
53k coqutee

191. Forgotten — ₱55
Verdant Modern Manor 74k
https://pastebin.com/LvnGg9qn

192. Molakkuma — ₱55
House Type / Cost: queenslander house / 53k
File name: https://pastebin.com/6TLt4ZNq

193. Paris — ₱55
House Type / Cost: tiny house / 71.4k
File name: parisg

194. Rose Gallery — ₱55
queenslander like
queenslander 65k

195. Workshop — ₱55
House Type / Cost: tiny house / 64.4kk
File name: workshop
```

The bot detects the price section, parses each numbered house, updates the matching price section in `Sheet1`, saves an `.xlsx`, and sends it back.

## Supported information

The parser recognizes:
- `₱55`
- `55 peso houses`
- mathematical/bold digits such as `𝟱𝟱`
- Pastebin/other URLs
- `File name: ...`
- `something file name`
- `House Type / Cost: queenslander house / 53k`
- `52k queenslander`
- `tiny house`
- `tiny home`
- `gingerbread house`
- `bunker house`
- `family house/home`
- `treehouse`
- `estate house`
- `mountain house`
- `racetrack house`
- `sandbox island`
- `hollywood house`
- `pizzaplace house`
- `shop house`
- `safari house`
- `friendly house`

Unknown house types default to `Tiny Home` rather than putting cost text into the House Types column.

## Important

The bot modifies a copy of the template. It does not overwrite `templates/template.xlsx`.

Generated files are stored in:

```text
output/
```
