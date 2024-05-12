import pandas as pd
import os

minwon_a = pd.read_excel("minwon_data.xlsx", parse_dates=["신청일시", "접수일시", "처리일자"])

#minwon_a.pivot_table(
#	index = "담당부서_3차분류",
#	columns = "처리일자",
#	values = "처리일자",
#	aggfunc = "count"minwon_delayed
#)

end_date = '2024-04-30'
until_m4 = pd.to_datetime(end_date)
minwon_b = minwon_a[minwon_a['접수일시'] <= until_m4]

not_processed = minwon_b[minwon_b['처리일자'].isna()]

end_date = '2024-03-31'
until_m3 = pd.to_datetime(end_date)
minwon_c = minwon_a[minwon_a['접수일시'] <= until_m3]
minwon_c.info()

# 이월 : 이전달까지 접수하여 이번달에 처리 + 아직 처리 안된 건
processed_3_not = minwon_c[minwon_c['처리일자'].isna()] #3월까지 접수되어 아직 처리 안됨
processed_3_4 = minwon_c[minwon_c['처리일자'] > until_m3] #3월까지 접수되어 4월에 처리

# processed_3_4 = processed_3_4.fillna(0)
# processed_3_not = processed_3_not.fillna(0)

start_date = '2024-04-01'
end_date = '2024-05-01'
m3_start = pd.to_datetime(start_date)
m3_end = pd.to_datetime(end_date)
minwon_submitted_4 = minwon_a[minwon_a['접수일시'].between(left = m3_start, right = m3_end)]
not_processed_4 = minwon_submitted_4[minwon_submitted_4['처리일자'].isna()]
minwon_submitted_df_4 = minwon_submitted_4

minwon_aa = minwon_a[minwon_a['처리일자'].between(left = m3_start, right = m3_end)]
minwon_ab = minwon_aa[minwon_a['처리일자'] <= minwon_a['처리완료예정일']]
 
minwon_delayed_processed_4 = minwon_aa[minwon_a['처리일자'] > minwon_a['처리완료예정일']]

minwon_4 = minwon_a.pivot_table(
    index = ["담당부서_1차분류", "담당부서_2차분류"],
    columns = "접수부서",
    values = "접수번호",
    aggfunc = "count",
    margins = True,
    margins_name = "Total"
)

# minwon_submitted_3 = minwon_submitted_3.pivot_table(
	# index = ["담당부서_1차분류", "담당부서_2차분류"],
    # columns = "접수부서",
    # values = "접수번호",
    # aggfunc = "count",
    # margins = True,
    # margins_name = "Total"
# )

minwon_submitted_4 = minwon_submitted_4.pivot_table(
    index = "담당부서_1차분류",
    columns = "합계부서",
    values = "접수번호",
    aggfunc = "count",
    # margins = True,
    # margins_name = "Total"
)

minwon_submitted_4 = minwon_submitted_4.stack()
minwon_submitted_4.info()

not_processed = not_processed.pivot_table(
    index = "담당부서_1차분류",
    columns = "합계부서",
    values = "접수번호",
    aggfunc = "count",
    # margins = True,
    # margins_name = "Total"
)
not_processed = not_processed.stack()

processed_3_4 = processed_3_4.pivot_table(
    index = "담당부서_1차분류",
    columns = "합계부서",
    values = "접수번호",
    aggfunc = "count",
    # margins = True,
    # margins_name = "Total"
)
processed_3_4 = processed_3_4.stack()

processed_3_not = processed_3_not.pivot_table(
    index = "담당부서_1차분류",
    columns = "합계부서",
    values = "접수번호",
    aggfunc = "count",
    # margins = True,
    # margins_name = "Total"
)
processed_3_not = processed_3_not.stack()

#delayed = not_processed_pre + not_processed_4
#delayed = processed_3_4 + processed_3_not
delayed = processed_3_4.add(processed_3_not, fill_value=0)

minwon_delayed_processed_4 = minwon_delayed_processed_4.pivot_table(
    index = "담당부서_1차분류",
    columns = "합계부서",
    values = "접수번호",
    aggfunc = "count",
    # margins = True,
    # margins_name = "Total"
)
minwon_delayed_processed_4 = minwon_delayed_processed_4.stack()

minwon_processed_normal = minwon_ab.pivot_table(
    index = "담당부서_1차분류",
    columns = "합계부서",
    values = "접수번호",
    aggfunc = "count",
    # margins = True,
    # margins_name = "Total"
)
minwon_processed_normal = minwon_processed_normal.stack()

df = pd.DataFrame({"이월": delayed, "4월접수": minwon_submitted_4, "정상":minwon_processed_normal, "지연처리":minwon_delayed_processed_4, "미처리": not_processed})
print(df)
df.info()

custom_order = ["중소기업정책실","창업벤처혁신실","소상공인정책실","기조실_대변인","운영지원과_감사실","지방청_국립공고"]
#df.index = df.index.set_levels(pd.Categorical(df.index.levels[1], categories=custom_order, ordered=True), level=1)
#df['합계부서'] = pd.Categorical(df['합계부서'], custom_order)
#df_sorted = df.sort_index(level=[1,0])
#df_sorted = df.sort_values("합계부서", ascending=False)
df = df.reset_index()

df['합계부서'] = pd.Categorical(df['합계부서'], custom_order)
df = df.sort_values('합계부서')
df = df.set_index(['담당부서_1차분류','합계부서'])

print(df)
#df = df.sort_values(by=["정상"], ascending=False)


# minwon_aa 4월에 처리된 민원
minwon_processed_4_satis = minwon_aa.pivot_table(
    index = ["담당부서_1차분류", "합계부서"],
    columns = "최종만족도",
    values = "만족도점수",
    aggfunc = "count",
    margins = True,
    margins_name = "Total"
)
minwon_processed_4_satis = minwon_processed_4_satis[['매우만족','만족','보통','불만','매우불만','평가없음',]]
minwon_processed_4_satis = minwon_processed_4_satis.reset_index()
minwon_processed_4_satis['합계부서'] = pd.Categorical(minwon_processed_4_satis['합계부서'], custom_order)
minwon_processed_4_satis = minwon_processed_4_satis.sort_values('합계부서')
minwon_processed_4_satis = minwon_processed_4_satis.set_index(['담당부서_1차분류','합계부서'])

minwon_processed_4_satis.fillna(0, inplace=True)

minwon_processed_4_satis["점수"] = minwon_processed_4_satis["매우만족"] * 100 + minwon_processed_4_satis["만족"] * 75 + minwon_processed_4_satis["보통"] * 50 + minwon_processed_4_satis["불만"] * 25

minwon_processed_4_satis["count"] = minwon_processed_4_satis["매우만족"] + minwon_processed_4_satis["만족"] + minwon_processed_4_satis["보통"] + minwon_processed_4_satis["불만"] + minwon_processed_4_satis["매우불만"]

minwon_processed_4_satis["만족도점수"] = minwon_processed_4_satis["점수"] / minwon_processed_4_satis["count"]

minwon_processed_4 = minwon_aa.pivot_table(
    index = ["담당부서_1차분류", "담당부서_2차분류"],
    columns = "접수부서",
    values = "처리일자",
    aggfunc = "count",
    margins = True,
    margins_name = "Total"
)

minwon_processed_4 = minwon_processed_4.stack()
#minwon_processed_4 = minwon_processed_4.sort_values(ascending=False)

#minwon_processed_3 = minwon_processed_3.melt(id_vars = "접수번호", value_vars="민원유형")

minwon_kind_3 = minwon_submitted_df_4.pivot_table(
    index = "민원종류",
    columns = "접수부서",
    values = "처리일자",
    aggfunc = "count",
    margins = True,
    margins_name = "Total"
)

satis_detail = minwon_aa.pivot_table(
    index = ["detail_grp1", "detail_grp2"],
    columns = "최종만족도",
    values = "만족도점수",
    aggfunc = "count",
    margins = True,
    margins_name = "Total"
)
satis_detail = satis_detail[['매우만족','만족','보통','불만','매우불만','평가없음',]]
# satis_detail = satis_detail.reset_index()
# satis_detail['합계부서'] = pd.Categorical(satis_detail['합계부서'], custom_order)
# satis_detail = satis_detail.sort_values('합계부서')
# satis_detail = satis_detail.set_index(['담당부서_1차분류','합계부서'])

satis_detail.fillna(0, inplace=True)
satis_detail["점수"] = satis_detail["매우만족"] * 100 + satis_detail["만족"] * 75 + satis_detail["보통"] * 50 + satis_detail["불만"] * 25 
satis_detail["count"] = satis_detail["매우만족"] + satis_detail["만족"] + satis_detail["보통"] + satis_detail["불만"] + satis_detail["매우불만"]
satis_detail["만족도점수"] = satis_detail["점수"] / satis_detail["count"]

processed_4_detail = minwon_aa.pivot_table(
    index = ["detail_grp1", "detail_grp2"],
    columns = "접수부서",
    values = "처리일자",
    aggfunc = "count",
    margins = True,
    margins_name = "Total"
)


#minwon_b = minwon_b.stack()

# if os.path.exists("monthly_report.xlsx"):
    # os.remove("monthly_report.xlsx")
# minwon_b.to_excel('monthly_report.xlsx')

# if os.path.exists("monthly_report.xlsx"):
    # os.remove("monthly_report.xlsx")
# minwon_submitted_3.to_excel('monthly_report.xlsx', sheet_name="3월접수")

with pd.ExcelWriter("monthly_report.xlsx") as writer:
#    minwon_b.to_excel(writer, sheet_name="3월까지 접수")
    minwon_4.to_excel(writer, sheet_name="4월까지 접수")
#    minwon_submitted_3.to_excel(writer, sheet_name="4월 접수") 
#    not_processed_pre.to_excel(writer, sheet_name="이월")
    df.to_excel(writer, sheet_name="접수현황")
    minwon_delayed_processed_4.to_excel(writer, sheet_name="지연처리")
    minwon_processed_4.to_excel(writer, sheet_name="처리현황상세")
    minwon_processed_4_satis.to_excel(writer, sheet_name="만족도")
    satis_detail.to_excel(writer, sheet_name="만족도상세")
    processed_4_detail.to_excel(writer, sheet_name="ccc")
    minwon_kind_3.to_excel(writer, sheet_name="민원종류")
