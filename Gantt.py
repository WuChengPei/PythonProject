import matplotlib.pyplot as plt
from datetime import datetime, timedelta

# 專案名稱2
num = int(input("你要輸入幾個專案?"))
projects = []

for i in range(num):
    project_name = input("依序輸入專案名稱:")
    projects.append(project_name)

start_dates = []
end_dates = []

print("請輸入各專案的開始和結束日期 (格式: YYYY-MM-DD)")
print("-" * 50)

for i, project in enumerate(projects):
    print(f"\n{project}:")
    
    # 輸入開始日期
    while True:
        try:
            start_input = input(f"  開始日期: ")
            start_date = datetime.strptime(start_input, "%Y-%m-%d")
            start_dates.append(start_date)
            break
        except ValueError:
            print("  格式錯誤，請使用 YYYY-MM-DD 格式 (例: 2023-09-13)")
    
    # 輸入結束日期
    while True:
        try:
            end_input = input(f"  結束日期: ")
            end_date = datetime.strptime(end_input, "%Y-%m-%d")
            
            # 檢查結束日期是否在開始日期之後
            if end_date <= start_date:
                print("  結束日期必須在開始日期之後，請重新輸入")
                continue
                
            end_dates.append(end_date)
            break
        except ValueError:
            print("  格式錯誤，請使用 YYYY-MM-DD 格式 (例: 2023-10-01)")

# 顏色
color = ['rosybrown', 'lightcoral', 'indianred', 'brown', 'maroon', 'sienna', 'chocolate', 'peru'] 
colors = []
for i in range(len(projects)):
    colors.append(color[i % 8])

# 繪製甘特圖
fig, ax = plt.subplots(figsize=(12, 5))

for i, project in enumerate(projects):
    start_date = start_dates[i]
    end_date = end_dates[i]
    duration = end_date - start_date

    # 使用自訂顏色繪製甘特圖
    ax.barh(project, left=start_date, width=duration, label=project, color=colors[i])

    # 新增結束日期
    ax.text(end_date, i, f'{end_date.strftime("%Y-%m-%d")}', va='center', ha='left', fontsize=10, color='black')

# 新增標題和x,y軸
ax.set_title('Gantt Chart')
ax.set_xlabel('dates')
ax.set_ylabel('projects')

# 設置X軸的日期刻度
date_format = '%Y-%m-%d'
ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter(date_format))
plt.xticks(rotation=45)  # 日期傾斜45度

ax.legend(loc='upper left')  # 新增圖例

# 顯示甘特圖
plt.tight_layout()
plt.show()