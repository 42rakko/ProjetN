import discord
from discord.ext import commands
from discord import app_commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
from dotenv import load_dotenv
import re


# .envファイルを読み込む
load_dotenv()
# 環境変数を取得
DISCORD_TOKEN = os.getenv('DISCORD_TOKEN')
GOOGLE_KEY_PATH = os.getenv('GOOGLE_KEY_PATH')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
SCHEDULE_SHEET = os.getenv('SCHEDULE_SHEET')
REQUEST_SHEET = os.getenv('REQUEST_SHEET')  
EXCHANGE_SHEET = os.getenv('EXCHANGE_SHEET')
PROXY_SHEET = os.getenv('PROXY_SHEET')
FEEDBACK_SHEET = os.getenv('FEEDBACK_SHEET')
STUDENT_SHEET = os.getenv('STUDENT_SHEET')
DISCORD_USER_SHEET = os.getenv('DISCORD_USER_SHEET')
# PUBLIC_SPREADSHEET_ID = os.getenv('PUBLIC_SPREADSHEET_ID')
# PUBLIC_SCHEDULE_SHEET = os.getenv('PUBLIC_SCHEDULE_SHEET')
DISCORD_SERVER_ID=691903146909237289
REQUEST_CHANNEL_ID=1227097579347509249
CONFIRM_CHANNEL_ID=1227910844990361640
FEEDBACK_CHANNEL_ID=1227097642610200586

# Google Sheets APIの認証設定
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_KEY_PATH, scope)
gspreadClient = gspread.authorize(creds)

# スプレッドシートのIDを指定
spreadsheet_id = SPREADSHEET_ID
# 掃除当番が記載されたシート名
schedule_sheet = SCHEDULE_SHEET
# 交換・代行の依頼を保存するシート名
request_sheet = REQUEST_SHEET
# 交換が成立したときに保存されるシート名
exchange_sheet = EXCHANGE_SHEET
# 代行が成立したときに保存されるシート名
proxy_sheet = PROXY_SHEET
# フィードバックが保存されるシート名
feedback_sheet = FEEDBACK_SHEET
# すべてのintra名が記載されているシート
student_sheet = STUDENT_SHEET
# discordのuser名とintra名をmapしているシート
discord_sheet = DISCORD_USER_SHEET


# # 公開用スプレッドシートのID
# public_spreadsheet_id = PUBLIC_SPREADSHEET_ID
# # 公開用スプレッドシートの掃除当番が記載されたシート名
# public_schedule_sheet = PUBLIC_SCHEDULE_SHEET

intents = discord.Intents.default()
intents.guilds = True
intents.members = True  # ここがTrueになっていることを確認
intents.message_content = True
discordClient = discord.Client(intents=intents)
tree = app_commands.CommandTree(discordClient)

# 特定のチャンネルのみでコマンドを実行できるようにするため
# チャンネル名が"command"のときにのみ実行できるデコレーター
def is_command_channel():
    async def predicate(ctx):
        return ctx.channel.name == "command"  # チャンネル名を"command"に設定
    return commands.check(predicate)

# MM-DD形式の日付に適切なYYYY-を付加して返す
def add_year(input_date: str) -> str:
    # 現在の日付を取得
    today = datetime.today()
    current_year = today.year
    current_month = today.month
    input_month = 0

    #前の２桁（月）をとる
    match = re.match(r"^\d{2}", input_date)
    if match:
        input_month = int(match.group())
    if input_month > 12 or input_month < 1:
        return ("")
    # 年を調整
    if current_month >= 9 and 1 <= input_month <= 4:
        # 当日が9-12月で、入力が1-4月 → 翌年
        target_year = current_year + 1
    elif current_month <= 4 and 9 <= input_month <= 12:
        # 当日が1-4月で、入力が9-12月 → 前年
        target_year = current_year - 1
    else:
        # それ以外 → 当年
        target_year = current_year
    # yyyy-mm-dd形式で返す
    return f"{target_year}-{input_date}"

# def copy2public():
#     source_sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(schedule_sheet)
#     target_sheet = gspreadClient.open_by_key(public_spreadsheet_id).worksheet(public_schedule_sheet)   
#     # コピー元の1列目と2列目のデータを取得
#     columns_to_copy = source_sheet.get_values('A:B')  # A列（1列目）からB列（2列目）を取得    
#     # コピー先にデータを書き込む
#     target_sheet.update(range_name='A:B', values=columns_to_copy)
#     #await interaction.followup.send("OK")

# @bot.event
@discordClient.event
# @is_command_channel()  # デコレーターを追加
async def on_ready():
    print("on_ready")
    await tree.sync() #スラッシュコマンドを同期


@tree.command(name="help", description="　　 利用可能なコマンドリストを表示する")
async def help_command(interaction: discord.Interaction):
    help_text = "利用できるコマンド:\n\n"

    # ツリーからすべてのコマンドを取得し、説明付きでリスト化
    for command in tree.walk_commands():
        help_text += f"/{command.name}: {command.description}\n"

    # ヘルプメッセージをユーザーに送信
    await interaction.response.send_message(help_text, ephemeral=True)

# @tree.command(name="hello", description="say hello")
# async def hello(interaction: discord.Interaction):
#     print("hello")
#     await interaction.response.send_message("hello world!", ephemeral=True)


#掃除の担当日を表示するコマンド whenの実装
@tree.command(name="when", description=" 　  指定するintra名の掃除担当日を表示する")
@app_commands.describe(
    intra="intra名",
)
async def when(
    interaction: discord.Interaction,
    intra: str,
):
    await interaction.response.defer(ephemeral=True)
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(schedule_sheet)
    data = sheet.get_all_records() #各行にアクセスできるようにする
    found_value = ""
    for row in data:
        if intra in row['logins']:
            logins_list = row['logins'].split()
            if intra in logins_list:
                found_value = found_value + "**" + row['date'] + "**  " + row['logins'] + "\n"
    if found_value != "":
        await interaction.followup.send(f"{found_value}\n<http://bit.ly/3BbrHBs>", ephemeral=True)
    else:
        await interaction.followup.send("intra名が存在しないか担当のアサインがありません")



#指定日の掃除担当者を表示するコマンド whoの実装
@tree.command(name="who", description="　　 指定日の担当者を表示する　　　　　　　日付はMM-DD形式")
@app_commands.describe(
    date="日付 (MM-DD)",
)
async def who(
    interaction: discord.Interaction,
    date: str,
):
    await interaction.response.defer(ephemeral=True)
    try:
        pattern = r"^\d{2}-\d{2}$"
        if not bool(re.match(pattern, date)):
            await interaction.followup.send("日付はMM-DD形式で入力してください", ephemeral=True)
            return
        if any(char.isdigit() for char in date):
            date = add_year(date)
            if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
                raise ValueError
    except ValueError:
        await interaction.followup.send("不正な日付が入力されました", ephemeral=True)
        return
    
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(schedule_sheet)
    data = sheet.get_all_records() #各行にアクセスできるようにする
    
    found_value = ""
    for row in data:
        if date in row['date']:
            found_value = found_value + "**" + row['date'] + "**  " + row['logins'] + "\n"

    if found_value != "":
        await interaction.followup.send(f"{found_value}\n<http://bit.ly/3BbrHBs>", ephemeral=True)
    else:
        await interaction.followup.send("当日の担当はありません")


# @bot.command()
# @is_command_channel()  # デコレーターを追加
# async def request(ctx, login_id:str, request_type: str, request_date:str, details:str):
@tree.command(name="request", description="   交換・代行のリクエストをする　　　　　日付はMM-DD形式")
@app_commands.describe(
    date="日付 (MM-DD)",
    intra="intra名",
    gender="性別",
    type="リクエストの種類（交換または代行）",
    others="その他（交換希望日等）"
)
@app_commands.choices(
    type=[
        app_commands.Choice(name="交換", value="交換"),
        app_commands.Choice(name="代行", value="代行"),
        app_commands.Choice(name="交換または代行", value="交換または代行"),
    ],
    gender=[
        app_commands.Choice(name="男性", value="男性"),
        app_commands.Choice(name="女性", value="女性")
    ]
)
async def request(
    interaction: discord.Interaction, 
    date: str, 
    intra: str, 
    gender: str,
    type: str, 
    others: str,
):
    # チャンネル ID をチェック    
    if interaction.channel_id != REQUEST_CHANNEL_ID:
        await interaction.response.send_message(f"requestコマンドはhttps://discord.com/channels/{DISCORD_SERVER_ID}/{REQUEST_CHANNEL_ID}で実行してください", 
            ephemeral=True
        )
        return
    await interaction.response.defer(ephemeral=False)
    try:
        pattern = r"^\d{2}-\d{2}$"
        if not bool(re.match(pattern, date)):
            await interaction.followup.send("日付はMM-DD形式で入力してください", ephemeral=True)
            return
        if any(char.isdigit() for char in date):
            date = add_year(date)
            if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
                raise ValueError
    except ValueError:
        await interaction.followup.send("不正な日付が入力されました", ephemeral=True)
        return
    today = datetime.today().strftime("%Y-%m-%d")
    if date < today:
        await interaction.followup.send("過去の日付は登録できません", ephemeral=True)
        return
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(schedule_sheet)
    data = sheet.get_all_records()
    row_index = next(
        (index for index, row in enumerate(data) 
        if date == row['date'] and intra in row['logins'].split()), 
        None
    )
    if row_index is not None:
        sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(request_sheet)
        data = sheet.get_all_records()
        row_index = next(
            (index for index, row in enumerate(data) 
            if row['date'] == date and intra in row['logins'].split()), 
            None
        )
        message = await interaction.followup.send(f"名前: {intra}\n性別: {gender}\n日時: {date}\n希望: {type}\n{others}", ephemeral=False)
        # message = await interaction.followup.send(f"名前: {intra}\n性別: {gender}\n日時: {date}\n希望: {type}\nその他: {others}\n\n?o(⁰ꇴ⁰o)三(o⁰ꇴ⁰)o?", ephemeral=False)
        new_data = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), date, type, intra, gender, others, str(message.id)]
        if row_index is not None:
            cell_range = f"A{row_index + 2}:G{row_index + 2}"
            sheet.update([new_data], cell_range)
        else:
            sheet.append_row(new_data)
            sheet.sort((2, 'asc'), range="A2:G1000")  # ２行目から2列目を基準に昇順にソートします
    else:
        await interaction.followup.send("日付またはintra名が誤っています", ephemeral=True)


@tree.command(name="rm", description="　   　リクエストを削除する　　　　　　　　　日付はMM-DD形式")
@app_commands.describe(
    date="日付 (MM-DD)",
    intra="intra名",
)
async def rm(
    interaction: discord.Interaction, 
    date: str, 
    intra: str, 
):
    # チャンネル ID をチェック
    if interaction.channel_id != REQUEST_CHANNEL_ID:
        await interaction.response.send_message(f"rmコマンドはhttps://discord.com/channels/{DISCORD_SERVER_ID}/{REQUEST_CHANNEL_ID}で実行してください", 
            ephemeral=True
        )
        return
    await interaction.response.defer(ephemeral=True)
    try:
        pattern = r"^\d{2}-\d{2}$"
        if not bool(re.match(pattern, date)):
            await interaction.followup.send("日付はMM-DD形式で入力してください", ephemeral=True)
            return
        if any(char.isdigit() for char in date):
            date = add_year(date)
            if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
                raise ValueError
    except ValueError:
        await interaction.followup.send("不正な日付が入力されました", ephemeral=True)
        return
    
    today = datetime.today().strftime("%Y-%m-%d")
    if date < today:
        await interaction.followup.send("過去の日付は登録できません", ephemeral=True)
        return
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(request_sheet)
    data = sheet.get_all_records()
    row_index = next(
        (index for index, row in enumerate(data) 
        if row['date'] == date and intra in row['logins'].split()), 
        None
    )
    if row_index is not None:
        sheet.delete_rows(row_index + 2)
        await interaction.followup.send("リクエストを削除しました", ephemeral=True)
    else:
        await interaction.followup.send("そのリクエストは存在しません", ephemeral=True)        



@tree.command(name="ls", description="　　　 募集中の交換・代行リストを表示する")
@app_commands.describe(
    gender="性別"
)
@app_commands.choices(
    gender=[
        app_commands.Choice(name="男性", value="男性"),
        app_commands.Choice(name="女性", value="女性")
    ]
)
async def list(
    interaction: discord.Interaction,
    gender: str,
):
    await interaction.response.defer(ephemeral=True)
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(request_sheet)
    data = sheet.get_all_records()
    if not data:
        await interaction.followup.send("募集中のリクエストはありません", ephemeral=True)
        return
    messages = []
    for row in data:
        if (gender == "男性"): 
            if row['gender'] == "男性":
                messages.append(
                    f"```"
                    f"日付: {row['date']}\n"
                    f"名前: {row['logins']}\n"
                    f"性別: {row['gender']}\n"
                    f"希望: {row['type']}\n"
                    f"{row['others']}\n"
                    f"```"
                    f"https://discord.com/channels/{DISCORD_SERVER_ID}/{REQUEST_CHANNEL_ID}/{row['message_id']}"
                )
        else:
            if row['gender'] == "女性":
                messages.append(
                    f"```"
                    f"日付: {row['date']}\n"
                    f"名前: {row['logins']}\n"
                    f"性別: {row['gender']}\n"
                    f"希望: {row['type']}\n"
                    f"{row['others']}\n"
                    f"```"
                    f"https://discord.com/channels/{DISCORD_SERVER_ID}/{REQUEST_CHANNEL_ID}/{row['message_id']}"
                )
    if not messages:
        await interaction.followup.send("募集中のリクエストはありません", ephemeral=True)
        return

    # 各行のデータをまとめ、コードブロックで囲む
    final_message = "\n" + "\n\n".join(messages) + "\n"
    await interaction.followup.send(final_message, ephemeral=True)


@tree.command(name="exchange", description="交換の成立をスケジュールに反映させる　日付はMM-DD形式")
@app_commands.describe(
    date1="1人目の日付 (MM-DD)",
    intra1="1人目のintra名",
    date2="2人目の日付 (MM-DD)",
    intra2="2人目のintra名",
)
async def exchange(
    interaction: discord.Interaction,
    date1: str,
    intra1: str,
    date2: str,
    intra2: str,
):
    # チャンネル ID をチェック
    if interaction.channel_id != CONFIRM_CHANNEL_ID:
        await interaction.response.send_message(f"exchangeコマンドはhttps://discord.com/channels/{DISCORD_SERVER_ID}/{CONFIRM_CHANNEL_ID}で実行してください", 
            ephemeral=True
        )
        return
    await interaction.response.defer(ephemeral=False)
    try:
        pattern = r"^\d{2}-\d{2}$"
        if not bool(re.match(pattern, date1)):
            await interaction.followup.send("date1はMM-DD形式で入力してください", ephemeral=True)
            return
        if not bool(re.match(pattern, date2)):
            await interaction.followup.send("date2はMM-DD形式で入力してください", ephemeral=True)
            return
        if any(char.isdigit() for char in date1):
            date1 = add_year(date1)
            if date1 != datetime.strptime(date1, "%Y-%m-%d").strftime("%Y-%m-%d"):
                raise ValueError
        if any(char.isdigit() for char in date2):
            date2 = add_year(date2)
            if date2 != datetime.strptime(date2, "%Y-%m-%d").strftime("%Y-%m-%d"):
                raise ValueError        
    except ValueError:
        await interaction.followup.send("不正な日付が入力されました", ephemeral=True)
        return

    today = datetime.today().strftime("%Y-%m-%d")
    if date1 < today or date2 < today:
        await interaction.followup.send("過去の日付は対応できません", ephemeral=True)
        return
    if date1 == date2:
        await interaction.followup.send("同一の日付は対応できません", ephemeral=True)
        return
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(schedule_sheet)
    data = sheet.get_all_records() #各行にアクセスできるようにする
    row1_index = next(
        (index for index, row in enumerate(data) 
        if row['date'] == date1 and intra1 in row['logins'].split()), 
        None
    )
    row2_index = next(
        (index for index, row in enumerate(data) 
        if row['date'] == date2 and intra2 in row['logins'].split()), 
        None
    )
    if row1_index is not None and row2_index is not None:
        row1_logins = data[row1_index]['logins'].replace(intra1, intra2)
        row2_logins = data[row2_index]['logins'].replace(intra2, intra1)  
        sheet.update_cell(row1_index + 2, 2, row1_logins)
        sheet.update_cell(row2_index + 2, 2, row2_logins)
        sheet_request = gspreadClient.open_by_key(spreadsheet_id).worksheet(request_sheet)
        data_request = sheet_request.get_all_records()
        row_index_request = next((index for index, row in enumerate(data_request) if row['date'] == date1 and row['logins'] == intra1), None)
        if row_index_request is not None:
            sheet_request.delete_rows(row_index_request + 2)
        sheet_request = gspreadClient.open_by_key(spreadsheet_id).worksheet(request_sheet)
        data_request = sheet_request.get_all_records()
        row_index_request = next((index for index, row in enumerate(data_request) if row['date'] == date2 and row['logins'] == intra2), None)
        if row_index_request is not None:
            sheet_request.delete_rows(row_index_request + 2)
        sheet_exchange = gspreadClient.open_by_key(SPREADSHEET_ID).worksheet(EXCHANGE_SHEET)
        new_data = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), date1, intra1, date2, intra2]
        sheet_exchange.append_row(new_data)
        # copy2public()

        #ここからmention用の文字列を作成する
        sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(discord_sheet)
        # 全データを取得
        data = sheet.get_all_values()  # 全てのデータを2次元リストとして取得
        # 2〜4列目に文字列が存在するか確認し、1列目の値を取得
        result1 = []
        for row in data:
            # 2〜4列目のどこかに文字列があるか確認
            if intra1 in row[1:4]:  # 1列目は `row[0]`、2列目以降が `row[1:]`
                result1.append(row[0])  # 同じ行の1列目を追加
                break
        result2 = []
        for row in data:
            # 2〜4列目のどこかに文字列があるか確認
            if intra2 in row[1:4]:  # 1列目は `row[0]`、2列目以降が `row[1:]`
                result2.append(row[0])  # 同じ行の1列目を追加
                break
        mention1 = ""
        if result1:
            mention1 = "<@" + result1[0] + ">"
        else:
            mention1 = intra1
        mention2 = ""
        if result2: 
            mention2 = "<@" + result2[0] + ">"
        else:
            mention2 = intra2
        #結果を出力する
        await interaction.followup.send(f"{date1} {intra1} <-> {date2} {intra2}\n\n {mention1} {mention2}\n5分程度たってから反映を確認してください\n<http://bit.ly/3BbrHBs>", ephemeral=False)
        # await interaction.followup.send(f"{date1} {intra1} <-> {date2} {intra2}\nヽ(\\*·ᗜ·)ﾉヽ(·ᗜ·\\* )ﾉ\n\n {mention1} {mention2}\n5分程度たってから反映を確認してください\nhttp://bit.ly/3BbrHBs", ephemeral=False)
    else:
        await interaction.followup.send("日付またはintra名が誤っています", ephemeral=True)


@tree.command(name="proxy", description=" 　  代行の成立をスケジュールに反映させる　日付はMM-DD形式")
@app_commands.describe(
    date="日付 (MM-DD)",
    intra1="代行してもらう人のintra名",
    intra2="代行する人のintra名"
)
async def proxy(
    interaction: discord.Interaction,
    date: str,
    intra1: str,
    intra2: str
):
    # チャンネル ID をチェック
    if interaction.channel_id != CONFIRM_CHANNEL_ID:
        await interaction.response.send_message(f"proxyコマンドはhttps://discord.com/channels/{DISCORD_SERVER_ID}/{CONFIRM_CHANNEL_ID}で実行してください", 
            ephemeral=True
        )
        return
    await interaction.response.defer(ephemeral=False)
    try:
        pattern = r"^\d{2}-\d{2}$"
        if not bool(re.match(pattern, date)):
            await interaction.followup.send("日付はMM-DD形式で入力してください", ephemeral=True)
            return
        if any(char.isdigit() for char in date):
            date = add_year(date)
            if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
                raise ValueError
    except ValueError:
        await interaction.followup.send("不正な日付が入力されました", ephemeral=True)
        return
    
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(student_sheet)
    students = sheet.col_values(1)
    if not intra2 in students:
        await interaction.followup.send("intra2が不正です", ephemeral=True)
        return

    today = datetime.today().strftime("%Y-%m-%d")
    if date < today:
        await interaction.followup.send("過去の日付は対応できません", ephemeral=True)
        return
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(schedule_sheet)
    data = sheet.get_all_records()
    row_index = next(
        (index for index, row in enumerate(data) 
        if row['date'] == date and intra1 in row['logins'].split()), 
        None
        )
    if row_index is not None:
        row_logins = data[row_index]['logins'].replace(intra1, intra2)
        sheet.update_cell(row_index + 2, 2, row_logins)
        sheet_request = gspreadClient.open_by_key(spreadsheet_id).worksheet(request_sheet)
        data_request = sheet_request.get_all_records()
        row_index_request = next((index for index, row in enumerate(data_request) if row['date'] == date and row['logins'] == intra1), None)
        if row_index_request is not None:
            sheet_request.delete_rows(row_index_request + 2)
        # copy2public()
        sheet_proxy = gspreadClient.open_by_key(spreadsheet_id).worksheet(proxy_sheet)
        new_data = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), date, intra1, intra2]
        sheet_proxy.append_row(new_data)        

        #ここからmention用の文字列を作成する
        sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(discord_sheet)
        # 全データを取得
        data = sheet.get_all_values()  # 全てのデータを2次元リストとして取得
        # 2〜4列目に文字列が存在するか確認し、1列目の値を取得
        result1 = []
        for row in data:
            # 2〜4列目のどこかに文字列があるか確認
            if intra1 in row[1:4]:  # 1列目は `row[0]`、2列目以降が `row[1:]`
                result1.append(row[0])  # 同じ行の1列目を追加
                break
        result2 = []
        for row in data:
            # 2〜4列目のどこかに文字列があるか確認
            if intra2 in row[1:4]:  # 1列目は `row[0]`、2列目以降が `row[1:]`
                result2.append(row[0])  # 同じ行の1列目を追加
                break
        mention1 = ""
        if result1:
            mention1 = "<@" + result1[0] + ">"
        else:
            mention1 = intra1
        mention2 = ""
        if result2: 
            mention2 = "<@" + result2[0] + ">"
        else:
            mention2 = intra2
        #結果を出力する
        await interaction.followup.send(f" {date} {intra1} -> {intra2}\n\n{mention1} {mention2}\n5分程度たってから反映を確認してください\n<http://bit.ly/3BbrHBs>", ephemeral=False)
        # await interaction.followup.send(f" {date} {intra1} -> {intra2}\n(¬_¬”)-cԅ(‾⌣‾ԅ)\n\n5分程度たってから反映を確認してください\n{mention1} {mention2}\nhttp://bit.ly/3BbrHBs", ephemeral=False)
    else:
        await interaction.followup.send("日付またはintra名が誤っています", ephemeral=True)


@tree.command(name="feedback", description="掃除の実施を報告する　　　　　　　　　日付はMM-DD形式（当日は'-'で省略可能）")
@app_commands.describe(
    date="日付 (MM-DD 当日は'-'で省略可能)",
    intras="intra名(複数人のときはspaceで区切る)",
    details="掃除箇所など自由記述"
)
async def feedback(
    interaction: discord.Interaction,
    date: str,
    intras: str,
    details: str,
):
    # チャンネル ID をチェック
    if interaction.channel_id != FEEDBACK_CHANNEL_ID:
        await interaction.response.send_message(f"feedbackコマンドはhttps://discord.com/channels/{DISCORD_SERVER_ID}/{FEEDBACK_CHANNEL_ID}で実行してください", 
            ephemeral=True
        )
        return
    await interaction.response.defer(ephemeral=False)  # 応答を準備
    if date == "-":
        date = datetime.today().strftime("%Y-%m-%d")    
    else:
        try:
            pattern = r"^\d{2}-\d{2}$"
            if not bool(re.match(pattern, date)):
                await interaction.followup.send("日付はMM-DD形式で入力してください", ephemeral=True)
                return
            if any(char.isdigit() for char in date):
                date = add_year(date)
                if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
                    raise ValueError
        except ValueError:
            await interaction.followup.send("不正な日付が入力されました", ephemeral=True)
            return
    
    today = datetime.today().strftime("%Y-%m-%d")
    if date > today:
        await interaction.followup.send("未来の日付は対応できません", ephemeral=True)
        return

    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(schedule_sheet)
    data = sheet.get_all_records() #各行にアクセスできるようにする
    
    row_index = next((index for index, row in enumerate(data) if row['date'] == date), None)

    if row_index is not None:
        intra_array = intras.split()
        write_value = ""
        found_value = ""
        none_value = ""
        # row_logins = data[row_index]['logins']

        st_sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(student_sheet)
        students = st_sheet.col_values(1)
        for intra in intra_array:
            # if intra in data[row_index]['logins']:
            #     found_value = found_value + intra + " " 
            #     if intra not in data[row_index]['feedback']:
            #         write_value = write_value + intra + " " 
            #     found_value = found_value + intra + " " 
            # else:
            #     none_value = none_value + intra + " "
            if not intra in students:
                none_value = none_value + intra + " "
            else:
                found_value = found_value + intra + " "
                if intra not in data[row_index]['feedback']:
                    write_value = write_value + intra + " "
        if write_value != "":
            feedback_data = data[row_index]['feedback']
            feedback_data = feedback_data + " " + write_value
            feedback_data = re.sub(r'\s{2,}', ' ', feedback_data)
            sheet.update_cell(row_index + 2, 3, feedback_data)
        if found_value != "":
            sheet_feedback = gspreadClient.open_by_key(spreadsheet_id).worksheet(feedback_sheet)
            new_data = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), date, found_value, details]
            sheet_feedback.append_row(new_data) 
            await interaction.followup.send(f"feedback:\n日付: {date} \nメンバー: {found_value}\n{details}", ephemeral=False)
        if none_value != "":
            await interaction.followup.send(f"intra名が誤っています: {none_value}", ephemeral=True)
    else:
        await interaction.followup.send("日付が誤っています", ephemeral=True)

discordClient.run(DISCORD_TOKEN)

