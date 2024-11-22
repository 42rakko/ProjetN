import discord
from discord.ext import commands
from discord import app_commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
from dotenv import load_dotenv


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
STUDENT_SHEET = os.getenv('STUDENT_SHEET')
DISCORD_USER_SHEET = os.getenv('DISCORD_USER_SHEET')
# PUBLIC_SPREADSHEET_ID = os.getenv('PUBLIC_SPREADSHEET_ID')
# PUBLIC_SCHEDULE_SHEET = os.getenv('PUBLIC_SCHEDULE_SHEET')


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


@tree.command(name="help", description="利用可能なコマンドリストを表示する")
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
@tree.command(name="when", description="指定するintra名の掃除担当日を表示する")
@app_commands.describe(
    intra="名前",
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
        await interaction.followup.send(f"{found_value}", ephemeral=False)
    else:
        await interaction.followup.send("intra名が存在しません")



#指定日の掃除担当者を表示するコマンド whoの実装
@tree.command(name="who", description="指定日の担当者を表示する 日付はYYYY-MM-DD形式")
@app_commands.describe(
    date="日付 (YYYY-MM-DD)",
)
async def who(
    interaction: discord.Interaction,
    date: str,
):
    await interaction.response.defer(ephemeral=True)
    try:
        if any(char.isdigit() for char in date):
            if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
                raise ValueError
    except ValueError:
        await interaction.followup.send("日付はYYYY-MM-DD形式で入力してください", ephemeral=True)
        return
    
    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(schedule_sheet)
    data = sheet.get_all_records() #各行にアクセスできるようにする
    
    found_value = ""
    for row in data:
        if date in row['date']:
            found_value = found_value + "**" + row['date'] + "**  " + row['logins'] + "\n"

    if found_value != "":
        await interaction.followup.send(f"{found_value}", ephemeral=False)
    else:
        await interaction.followup.send("日付が誤っています")




# @bot.command()
# @is_command_channel()  # デコレーターを追加
# async def request(ctx, login_id:str, request_type: str, request_date:str, details:str):
@tree.command(name="request", description="交換・代行のリクエストをする")
@app_commands.describe(
    intra="申請者の名前",
    type="リクエストの種類（交換または代行）",
    date="希望する日付 (YYYY-MM-DD)",
    gender="性別",
    others="その他（交換希望日等）"
)
@app_commands.choices(
    type=[
        app_commands.Choice(name="交換", value="交換"),
        app_commands.Choice(name="代行", value="代行"),
        app_commands.Choice(name="交換または代行", value="交換または代行")
    ],
    gender=[
        app_commands.Choice(name="男性", value="男性"),
        app_commands.Choice(name="女性", value="女性")
    ]
)
async def request(
    interaction: discord.Interaction, 
    intra: str, 
    type: str, 
    date: str, 
    gender: str,
    others: str,
):
    await interaction.response.defer(ephemeral=False)
    try:
        if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
            raise ValueError
    except ValueError:
        await interaction.followup.send("日付はYYYY-MM-DD形式で入力してください", ephemeral=True)
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
        new_data = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), date, type, intra, gender, others]
        if row_index is not None:
            cell_range = f"A{row_index + 2}:F{row_index + 2}"
            sheet.update([new_data], cell_range)
        else:
            sheet.append_row(new_data)
            sheet.sort((2, 'asc'), range="A2:F1000")  # ２行目から2列目を基準に昇順にソートします
        await interaction.followup.send(f"名前: {intra}\n性別: {gender}\n日時: {date}\n希望: {type}\nその他: {others}", ephemeral=False)
    else:
        await interaction.followup.send("日付またはintra名が誤っています", ephemeral=True)




@tree.command(name="ls", description="募集中の交換・代行のリクエストリストを表示する")
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
                    f"日付: {row['date']}\n"
                    f"名前: {row['logins']}\n"
                    f"性別: {row['gender']}\n"
                    f"希望: {row['type']}\n"
                    f"その他: {row['others']}\n"
                )
        else:
            if row['gender'] == "女性":
                messages.append(
                    f"日付: {row['date']}\n"
                    f"名前: {row['logins']}\n"
                    f"性別: {row['gender']}\n"
                    f"希望: {row['type']}\n"
                    f"その他: {row['others']}\n"
                )
    if not messages:
        await interaction.followup.send("募集中のリクエストはありません", ephemeral=True)
        return

    # 各行のデータをまとめ、コードブロックで囲む
    final_message = "```\n" + "\n\n".join(messages) + "\n```"
    await interaction.followup.send(final_message, ephemeral=True)





@tree.command(name="exchange", description="交換の成立をリストに反映させる")
@app_commands.describe(
    date1="１人目の日付",
    intra1="１人目の名前",
    date2="２人目の日付",
    intra2="２人目の名前",
)
async def exchange(
    interaction: discord.Interaction,
    date1: str,
    intra1: str,
    date2: str,
    intra2: str,
):
    await interaction.response.defer(ephemeral=False)
    try:
        if date1 != datetime.strptime(date1, "%Y-%m-%d").strftime("%Y-%m-%d") or \
           date2 != datetime.strptime(date2, "%Y-%m-%d").strftime("%Y-%m-%d"):
            raise ValueError
    except ValueError:
        await interaction.followup.send("日付はYYYY-MM-DD形式で入力してください", ephemeral=True)
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
        sheet_exchange = gspreadClient.open_by_key(spreadsheet_id).worksheet(exchange_sheet)
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
        await interaction.followup.send(f"{date1} {intra1} <-> {date2} {intra2}\n5分程度たってから反映を確認してください {mention1} {mention2}", ephemeral=False)
    else:
        await interaction.followup.send("日付またはintra名が誤っています", ephemeral=True)

@tree.command(name="proxy", description="代行の成立をリストに反映させる")
@app_commands.describe(
    date="日付",
    intra1="代行してもらう人の名前",
    intra2="代行する人の名前"
)
async def proxy(
    interaction: discord.Interaction,
    date: str,
    intra1: str,
    intra2: str
):
    await interaction.response.defer(ephemeral=False)
    try:
        if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
            raise ValueError
    except ValueError:
        await interaction.followup.send("日付はYYYY-MM-DD形式で入力してください", ephemeral=False)
        return

    sheet = gspreadClient.open_by_key(spreadsheet_id).worksheet(student_sheet)
    students = sheet.col_values(1)
    if not intra2 in students:
        await interaction.followup.send("intra2が不正です", ephemeral=False)
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
        await interaction.followup.send(f" {date} {intra1} -> {intra2}\n5分程度たってから反映を確認してください {mention1} {mention2}", ephemeral=False)
    else:
        await interaction.followup.send("日付またはintra名が誤っています", ephemeral=True)

@tree.command(name="feedback", description="掃除の実施を報告する 日付はYYYY-MM-DD形式")
@app_commands.describe(
    date="日付",
    intras="名前(複数人のときはspaceで区切る)",
    details="掃除箇所など自由記述"
)
async def feedback(
    interaction: discord.Interaction,
    date: str,
    intras: str,
    details: str,
):    
    await interaction.response.defer(ephemeral=False)  # 応答を準備
    try:
        if date != datetime.strptime(date, "%Y-%m-%d").strftime("%Y-%m-%d"):
            raise ValueError
    except ValueError:
        await interaction.followup.send("日付はYYYY-MM-DD形式で入力してください", ephemeral=False)
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
            sheet.update_cell(row_index + 2, 3, feedback_data + write_value)
        if found_value != "":
            await interaction.followup.send(f"feedback:\n日付: {date} \nメンバー: {found_value}\n掃除箇所: {details}", ephemeral=True)
        if none_value != "":
            await interaction.followup.send(f"intra名が誤っています: {none_value}", ephemeral=True)
    else:
        await interaction.followup.send("日付が誤っています", ephemeral=True)

discordClient.run(DISCORD_TOKEN)

