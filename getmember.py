import discord
from discord.ext import commands
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
import os

# .envファイルを読み込む
load_dotenv()

DISCORD_TOKEN = os.getenv('DISCORD_TOKEN')

intents = discord.Intents.default()
intents.guilds = True
intents.members = True  # ここがTrueになっていることを確認
intents.message_content = True

# bot = commands.Bot(command_prefix="!", intents=intents)
bot = commands.Bot(command_prefix="!", intents=intents)  # デフォルトのヘルプコマンドを無効にする

@bot.event
async def on_ready():
    await bot.change_presence(status=discord.Status.online)
    print(f'Logged in as {bot.user}')
    print("ボットが参加しているサーバー：")
    for guild in bot.guilds:
        print(f"- {guild.name} (ID: {guild.id})")
    # コマンドが使用されたサーバー（ギルド）を取得
    #guild = ctx.guild
    #await ctx.send(f'{member.id},{member.name}')  # チャットにも出力
    # ファイルを開いて書き込み開始
    with open("members.csv", "w", encoding="utf-8") as file:
        file.write("ID,Name,DisplayName,NickName\n")  # CSVのヘッダー行を書き出し

        # メンバー情報を非同期で取得し、ファイルに書き込み
        async for member in guild.fetch_members(limit=None):
            file.write(f"{member.id},{member.name},{member.display_name},{member.nick}\n")

    print("メンバーをmembers.csvに出力しました：Ctrl-c で終了してください")


@bot.command()
async def hello(ctx):
    print('Hello, world!')
    # await ctx.send('Hello, world!')

bot.run(DISCORD_TOKEN)

