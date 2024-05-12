import pandas as pd
import os
import math

minwon = pd.read_excel("민원현황목록_민원처리대장_20240501093330.xlsx", sheet_name = 0, header=[0,1])
#minwon_a = pd.read_excel("minwon_data.xlsx", parse_dates=["신청일시", "접수일시", "처리완료예정일", "처리일자"])
#minwon.fillna("", inplace=True)

c = []
for f1, f2 in minwon.columns:
    tmp = f'{f1}'
    if f2.startswith('Unnamed') == False:
       tmp = f'{f1}_{f2}'
    c.append(''.join(tmp.split()))

minwon.columns = c

minwon_a = minwon[['신청번호', '신청일시', '접수번호', '접수일시', '민원처리기한_설정된처리일', '민원처리기한_처리완료예정일', '민원처리기한_처리잔여일', '처리연장횟수', '처리일자', '처리기간', '담당부서_1차분류', '담당부서_2차분류', '담당부서_3차분류', '담당부서_4차분류', '민원종류', '최종만족도']]
minwon_a = minwon_a.set_index(keys = ["접수번호"])

minwon_a = minwon_a.rename(columns = {"민원처리기한_설정된처리일":"설정된처리일", "민원처리기한_처리완료예정일":"처리완료예정일", "민원처리기한_처리잔여일":"처리잔여일"})

minwon_a['최종만족도'].fillna('평가없음', inplace=True)

minwon_a.info()

minwon_a['접수일시'] = pd.to_datetime(minwon_a['접수일시'], errors='coerce')
minwon_a['처리일자'] = pd.to_datetime(minwon_a['처리일자'], errors='coerce')
minwon_a['처리완료예정일'] = pd.to_datetime(minwon_a['처리완료예정일'], errors='coerce')

end_date = '2024-04-30'
until_m3 = pd.to_datetime(end_date)
minwon_b = minwon_a[minwon_a['접수일시'] <= until_m3]
minwon_submitted = minwon_a[minwon_a['접수일시'] <= until_m3]

minwon_submitted = minwon_submitted[['접수일시', '담당부서_1차분류', '담당부서_2차분류', '담당부서_3차분류', '담당부서_4차분류']]
#minwon_submitted = minwon_submitted.set_index('접수번호')

minwon_submitted.info()

submit_division = pd.Series(index = minwon_submitted.index, name="접수부서")
#minwon_submitted.reset_index()
for i, row in minwon_submitted.iterrows():
    # print(f"index -> {i}")
    # print(f"row -> {row}")
    last_division = ""
    for j, val in row.items():
        #print(row.loc[j])
        #print(f"index {j} -> value >{val}<")
        division = val
        if pd.isna(division):
            #print(last_division)
            submit_division[i] = last_division
            last_division = ""
            break
        if j == "담당부서_4차분류":
            submit_division[i] = division
            last_division = ""
            break
        last_division = division

minwon_data = pd.concat([minwon_b, submit_division], axis = 1)

sum_division = pd.Series(index = minwon_data.index, name="합계부서")
region_division = ['경기북부사무소','경북북부사무소','소상공인과','성장지원과','지역정책과','지역혁신과','창업벤처과','인천지방중소벤처기업청']
etc1 = ['정보화담당관','홍보담당관','기획혁신담당관','재정행정담당관','기획조정실','디지털소통팀']
etc2 = ['감사담당관','옴부즈만지원단','운영지원과']

for i, row in minwon_data.iterrows():
    for j, val in row.items():
        if j == "담당부서_2차분류":
            division = val
        if j == "접수부서":
            if val in region_division:
                sum_division[i] = '지방청_국립공고'
            elif val in etc1:
                sum_division[i] = '기조실_대변인'
            elif val in etc2:
                sum_division[i] = '운영지원과_감사실'
            else:
                sum_division[i] = division
            
minwon_data = pd.concat([minwon_data, sum_division], axis = 1)

detail_grp1 = pd.Series(index = minwon_data.index, name="detail_grp1")
region_division = ['경기북부사무소','경북북부사무소','소상공인과','성장지원과','지역정책과','지역혁신과','창업벤처과','인천지방중소벤처기업청']
region_branch = ['경기지방중소벤처기업청','서울지방중소벤처기업청','충북지방중소벤처기업청','울산지방중소벤처기업청','부산지방중소벤처기업청','대구.경북지방중소벤처기업청','경남지방중소벤처기업청','전북지방중소벤처기업청','충남지방중소벤처기업청','광주.전남지방중소벤처기업청','인천지방중소벤처기업청','강원지방중소벤처기업청']
branch_abb = {"경기지방중소벤처기업청":"경기청","서울지방중소벤처기업청":"서울청","충북지방중소벤처기업청":"충북청","울산지방중소벤처기업청":"울산청","부산지방중소벤처기업청":"부산청","대구.경북지방중소벤처기업청":"대구경북청","경남지방중소벤처기업청":"경남청","전북지방중소벤처기업청":"전북청","충남지방중소벤처기업청":"충남청","광주.전남지방중소벤처기업청":"광주전남청","인천지방중소벤처기업청":"인청청","강원지방중소벤처기업청":"강원청"}

for i, row in minwon_data.iterrows():
    for j, val in row.items():
        if j == "담당부서_2차분류":
            if val in region_branch:
                detail_grp1[i] = '지방청'
            elif val == "기획조종실":
                detail_grp1[i] = '기조실'
            else:
                detail_grp1[i] = val
minwon_data = pd.concat([minwon_data, detail_grp1], axis = 1)

detail_grp2 = pd.Series(index = minwon_data.index, name="detail_grp2")
for i, row in minwon_data.iterrows():
    grp2 = ""
    for j, val in row.items():
        if j == "담당부서_2차분류" and val in region_branch:
            if val in branch_abb.keys():
                grp2 = branch_abb[val]
        if j == "접수부서":
            if grp2 == "":
                grp2 = val
    detail_grp2[i] = grp2
    grp2 = ""

minwon_data = pd.concat([minwon_data, detail_grp2], axis = 1)

#print(submit_division.to_string())
#submit_division.to_excel('submitted.xlsx', index = True )

def evaluate(val):
    v = 0
    if val == "매우만족": v = 100
    if val == "만족": v = 75
    if val == "보통": v = 50
    if val == "불만": v = 25
    if val == "매우불만": v = 0
    if val == "평가없음": v = 0
    
    return v

satis_point = pd.Series(index = minwon_data.index, name="만족도점수")
for i, row in minwon_data.iterrows():
    for j, val in row.items():
        if j == "최종만족도":
            r = evaluate(val)
            satis_point[i] = r

minwon_data = pd.concat([minwon_data, satis_point], axis = 1)

# process_in_period = pd.Series(index = minwon_b.index, name="기간내처리")
# for i, row in minwon_b.iterrows():
    # for j, val in row.items():
        # due_date = row["처리완료예정일"]
        # processed_date = row["처리일자"]
        # if processed_date <= due_date:
            # process_in_period[i] = 'y'
        # else:
            # process_in_period[i] = 'n'

# minwon_data = pd.concat([minwon_data, process_in_period], axis = 1)

if os.path.exists("monthly_data.xlsx"):
    os.remove("monthly_data.xlsx")
minwon_data.to_excel('minwon_data.xlsx', index = True)
        
