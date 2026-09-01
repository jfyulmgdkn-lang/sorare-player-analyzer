import os
import sys
import asyncio
import requests
import discord
from discord import app_commands
from discord.ext import commands
from dotenv import load_dotenv

API_URL = "https://api.sorare.com/graphql"
CURRENT_SEASON_YEAR = 2026

LEAGUES = {
    "Bundesliga": "bundesliga-de",
    "2. Bundesliga": "2-bundesliga",
    "Premier League": "premier-league-gb-eng",
    "La Liga": "laliga-es",
    "Ligue 1": "ligue-1-fr",
    "Ligue 2": "ligue-2-fr",
    "MLS": "mlspa",
    "Bundesliga Österreich": "austrian-bundesliga",
    "HNL": "1-hnl",
    "Primeira Liga": "primeira-liga-pt",
    "Jupiler Pro League": "jupiler-pro-league",
}

POSITIONS = {
    "Torwart": "Goalkeeper",
    "Abwehr": "Defender",
    "Mittelfeld": "Midfielder",
    "Sturm": "Forward",
}

BUDGETS = {
    "Bis 15 €": 15.0,
    "Bis 25 €": 25.0,
    "Bis 50 €": 50.0,
    "Preis egal": None,
}


def graphql(api_key, query, variables=None):
    try:
        response = requests.post(
            API_URL,
            headers={
                "content-type": "application/json",
                "APIKEY": api_key,
            },
            json={"query": query, "variables": variables or {}},
            timeout=45,
        )
    except requests.RequestException as error:
        return None, f"Verbindungsfehler: {error}"

    try:
        payload = response.json()
    except ValueError:
        return None, f"HTTP {response.status_code}: Ungültige Antwort"

    if payload.get("errors"):
        messages = [e.get("message", str(e)) for e in payload["errors"]]
        return None, " | ".join(messages)

    if response.status_code != 200:
        return None, f"HTTP {response.status_code}"

    return payload.get("data"), None


def avg(values):
    return sum(values) / len(values) if values else 0.0


def pct(part, whole):
    return (part / whole * 100.0) if whole else 0.0


def calculate_stats(scores):
    clean = [float(x) for x in scores if x is not None]

    if not clean:
        return None

    l5 = avg(clean[:5])
    l10 = avg(clean[:10])
    l15 = avg(clean[:15])

    count60 = sum(x >= 60 for x in clean)
    count75 = sum(x >= 75 for x in clean)
    count90 = sum(x >= 90 for x in clean)

    rate60 = pct(count60, len(clean))
    rate75 = pct(count75, len(clean))
    rate90 = pct(count90, len(clean))

    return {
        "l5": l5,
        "l10": l10,
        "l15": l15,
        "rate60": rate60,
        "rate75": rate75,
        "rate90": rate90,
        "games": len(clean),
    }


def fetch_league_players(api_key, league_slug):
    query = """
    query LeaguePlayers($leagueSlug: String!, $after: String) {
      football {
        competition(slug: $leagueSlug) {
          orderedPlayers(first: 100, after: $after) {
            nodes {
              slug
              displayName
              position
            }
            pageInfo {
              hasNextPage
              endCursor
            }
          }
        }
      }
    }
    """

    all_players = []
    after = None

    while True:
        data, error = graphql(
            api_key,
            query,
            {"leagueSlug": league_slug, "after": after},
        )

        if error:
            return None, error

        football = (data or {}).get("football") or {}
        competition = football.get("competition")

        if not competition:
            return None, "Liga wurde nicht gefunden."

        connection = competition.get("orderedPlayers") or {}
        all_players.extend(connection.get("nodes") or [])

        page_info = connection.get("pageInfo") or {}

        if not page_info.get("hasNextPage"):
            break

        after = page_info.get("endCursor")
        if not after:
            break

    unique = {}

    for player in all_players:
        slug = (player or {}).get("slug")
        if slug:
            unique[slug] = player

    return list(unique.values()), None


def fetch_player_scores(api_key, slug):
    query = """
    query PlayerScores($slug: String!) {
      anyPlayer(slug: $slug) {
        slug
        displayName
        ... on Player {
          position
          rawPlayerGameScores(last: 15, lowCoverage: true)
        }
      }
    }
    """

    data, error = graphql(api_key, query, {"slug": slug})

    if error:
        return None, error

    player = (data or {}).get("anyPlayer")

    if not player:
        return None, "Spieler nicht gefunden"

    stats = calculate_stats(player.get("rawPlayerGameScores") or [])

    if not stats:
        return None, "Keine Scores"

    return {
        "name": player.get("displayName") or slug,
        "slug": player.get("slug") or slug,
        "position": player.get("position"),
        "stats": stats,
    }, None


def fetch_lowest_inseason_limited_price(api_key, player_slug):
    query = """
    query LiveOffers($playerSlug: String!, $after: String) {
      tokens {
        liveSingleSaleOffers(
          first: 50
          after: $after
          playerSlug: $playerSlug
          sport: FOOTBALL
        ) {
          nodes {
            senderSide {
              anyCards {
                rarityTyped
                seasonYear
              }
            }
            receiverSide {
              amounts {
                eurCents
              }
            }
          }
          pageInfo {
            hasNextPage
            endCursor
          }
        }
      }
    }
    """

    after = None
    prices = []

    while True:
        data, error = graphql(
            api_key,
            query,
            {"playerSlug": player_slug, "after": after},
        )

        if error:
            return None, error

        offers = ((data or {}).get("tokens") or {}).get("liveSingleSaleOffers") or {}

        for offer in offers.get("nodes") or []:
            sender = (offer or {}).get("senderSide") or {}
            cards = sender.get("anyCards") or []

            valid = any(
                str((card or {}).get("rarityTyped") or "").lower() == "limited"
                and (card or {}).get("seasonYear") == CURRENT_SEASON_YEAR
                for card in cards
            )

            if not valid:
                continue

            amounts = ((offer or {}).get("receiverSide") or {}).get("amounts") or {}
            eur_cents = amounts.get("eurCents")

            if eur_cents is None:
                continue

            try:
                price = float(eur_cents) / 100.0
            except (TypeError, ValueError):
                continue

            if price > 0:
                prices.append(price)

        page_info = offers.get("pageInfo") or {}

        if not page_info.get("hasNextPage"):
            break

        after = page_info.get("endCursor")
        if not after:
            break

    return (min(prices) if prices else None), None


def run_analysis(api_key, league_name, position_name, budget_name):
    league_slug = LEAGUES[league_name]
    position_api = POSITIONS[position_name]
    budget = BUDGETS[budget_name]

    players, error = fetch_league_players(api_key, league_slug)
    if error:
        raise RuntimeError(error)

    position_players = [
        player for player in players
        if player.get("position") == position_api
    ]

    results = []

    for player in position_players:
        slug = player.get("slug")

        result, error = fetch_player_scores(api_key, slug)
        if error or not result:
            continue

        price, _ = fetch_lowest_inseason_limited_price(api_key, slug)
        result["price"] = price

        if budget is not None:
            if price is None or price > budget:
                continue

        results.append(result)

    results.sort(
        key=lambda item: item["stats"]["l15"],
        reverse=True,
    )

    return results[:20]


def create_embed(results, league_name, position_name, budget_name):
    embed = discord.Embed(
        title=f"🏆 Top 20 nach L15 – {league_name} – {position_name}",
        description=(
            f"**Budget:** {budget_name}\n"
            f"**Karten:** Limited In-Season {CURRENT_SEASON_YEAR}/{str(CURRENT_SEASON_YEAR + 1)[-2:]}\n"
            f"**Spiele:** Verein + Nationalmannschaft"
        ),
    )

    if not results:
        embed.add_field(
            name="Keine Ergebnisse",
            value="Für diese Auswahl wurden keine passenden Spieler gefunden.",
            inline=False,
        )
        return embed

    lines = []

    for index, player in enumerate(results, start=1):
        s = player["stats"]
        price = player.get("price")
        price_text = "–" if price is None else f"{price:.2f} €"

        lines.append(
            f"**{index}. {player['name']}**\n"
            f"**L15 {s['l15']:.1f}** | L5 {s['l5']:.1f} | L10 {s['l10']:.1f} | "
            f"60+ {s['rate60']:.0f}% | 75+ {s['rate75']:.0f}% | 90+ {s['rate90']:.0f}% | "
            f"💶 {price_text}"
        )

    # Discord Embed-Felder dürfen nur begrenzt lang sein.
    chunks = []
    current = ""

    for entry in lines:
        candidate = current + ("\n\n" if current else "") + entry

        if len(candidate) > 1000:
            chunks.append(current)
            current = entry
        else:
            current = candidate

    if current:
        chunks.append(current)

    for i, chunk in enumerate(chunks, start=1):
        embed.add_field(
            name="Ranking" if i == 1 else f"Ranking – Teil {i}",
            value=chunk,
            inline=False,
        )

    return embed


load_dotenv()

DISCORD_TOKEN = os.getenv("DISCORD_TOKEN", "").strip()
SORARE_API_KEY = os.getenv("SORARE_API_KEY", "").strip()
GUILD_ID_RAW = os.getenv("GUILD_ID", "").strip()

if not DISCORD_TOKEN:
    print("❌ DISCORD_TOKEN fehlt in der .env-Datei.")
    sys.exit(1)

if not SORARE_API_KEY:
    print("❌ SORARE_API_KEY fehlt in der .env-Datei.")
    sys.exit(1)

GUILD_ID = int(GUILD_ID_RAW) if GUILD_ID_RAW.isdigit() else None

intents = discord.Intents.default()
bot = commands.Bot(command_prefix="!", intents=intents)


@bot.event
async def on_ready():
    print(f"✅ Bot eingeloggt als: {bot.user}")

    try:
        if GUILD_ID:
            guild = discord.Object(id=GUILD_ID)
            bot.tree.copy_global_to(guild=guild)
            synced = await bot.tree.sync(guild=guild)
            print(f"✅ {len(synced)} Befehle auf Testserver synchronisiert.")
        else:
            synced = await bot.tree.sync()
            print(f"✅ {len(synced)} globale Befehle synchronisiert.")
    except Exception as error:
        print(f"❌ Fehler beim Synchronisieren: {error}")


@app_commands.command(
    name="analyse",
    description="Zeigt die Top 20 Sorare-Spieler nach Liga, Position und Budget."
)
@app_commands.choices(
    liga=[
        app_commands.Choice(name=name, value=name)
        for name in LEAGUES.keys()
    ],
    position=[
        app_commands.Choice(name=name, value=name)
        for name in POSITIONS.keys()
    ],
    budget=[
        app_commands.Choice(name=name, value=name)
        for name in BUDGETS.keys()
    ],
)
async def analyse(
    interaction: discord.Interaction,
    liga: app_commands.Choice[str],
    position: app_commands.Choice[str],
    budget: app_commands.Choice[str],
):
    await interaction.response.defer(thinking=True)

    try:
        results = await asyncio.to_thread(
            run_analysis,
            SORARE_API_KEY,
            liga.value,
            position.value,
            budget.value,
        )

        embed = create_embed(
            results,
            liga.value,
            position.value,
            budget.value,
        )

        await interaction.followup.send(embed=embed)

    except Exception as error:
        await interaction.followup.send(
            f"❌ Fehler bei der Analyse:\n```{str(error)[:1500]}```"
        )


bot.tree.add_command(analyse)

bot.run(DISCORD_TOKEN)
