Poke_Num_List = []
Poke_Name_List = []


Poke_List_URL = 'https://wiki.ポケモン.com/wiki/ポケモン一覧'
Poke_Type_List = ['くさ','どく','ほのお','ひこう','みず','むし','ノーマル','あく','でんき','エスパー','じめん','こおり','はがね','フェアリー','かくとう','いわ','ドラゴン']

Poke_Except_List = Poke_Type_List + ['全国ナンバー','タイプ','アローラのすがた','ガラルのすがた','ヒスイのすがた','パルデアのすがた','たいようのすがた','あまみずのすがた','ゆきぐものすがた',
                    'ゲンシカイオーガ','ゲンシグラードン','くさきのミノ','すなちのミノ','ゴミのミノ','ヒートロトム','ウォッシュロトム','スピンロトム','フロストロトム','カットロトム',
                    'ランドフォルム','スカイフォルム','ダルマモード','ボイスフォルム','ステップフォルム','いましめられしフーパ','ときはなたれしフーパ','めらめらスタイル','ぱちぱちスタイル',
                    'ふらふらスタイル','まいまいスタイル','たそがれのたてがみ','あかつきのつばさ','ウルトラネクロズマ','れきせんのゆうしゃ','けんのおう','たてのおう','いちげきのかた',
                    'れんげきのかた','はくばじょうのすがた','こくばじょうのすがた']

LogFile = 'Log.txt'
# ExcelFile = 'PokeData.xlsx'

SmallChara_List = ['ァ','ィ','ゥ','ェ','ォ','ャ','ュ','ョ']
BigChara_List = ['ア','イ','ウ','エ','オ','ヤ','ユ','ヨ']

Bef_Special_List = ['♀','♂','2','Z']
Aft_Special_List = ['メス','オス','ツー','ゼット']

Kana = 'カキクケコサシスセソタチツテトハヒフヘホハヒフヘホ'
DakuKana = 'ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポ'

Hyphen_JD = True #TRUE：直前の文字,#FALSE：伸ばした母音 (カイリュー等最後がハイフンで終わる場合に、直前の文字を使うか、伸ばした母音を使うか)
Dakuten_JD = True #TRUE：濁点を除去、#FALSE：濁点を除去しない

if Hyphen_JD and Dakuten_JD:
    ExcelFile = '01_PokeData_Tyokuzen_NoDakuten.xlsx'
elif Hyphen_JD == False and Dakuten_JD:
    ExcelFile = '02_PokeData_TyokuBoin_NoDakuten.xlsx'
elif Hyphen_JD and Dakuten_JD == False:
    ExcelFile = '03_PokeData_Tyokuzen_YesDakuten.xlsx'
elif Hyphen_JD == False and Dakuten_JD == False:
    ExcelFile = '04_PokeData_TyokuBoin_YesDakuten.xlsx'