"""Generate sample figures for demo."""
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.family'] = 'Times New Roman'

# Figure 1: bar chart - AI adoption by industry
fig, ax = plt.subplots(figsize=(6, 3.5))
industries = ['Tài chính', 'Bán lẻ', 'Sản xuất', 'Y tế', 'Giáo dục']
adoption = [68, 54, 47, 39, 28]
ax.bar(industries, adoption, color='#1f4e79')
ax.set_ylabel('Tỷ lệ ứng dụng (%)')
ax.set_title('Tỷ lệ ứng dụng AI theo ngành tại Việt Nam (2025)')
ax.set_ylim(0, 80)
for i, v in enumerate(adoption):
    ax.text(i, v + 1.5, f'{v}%', ha='center', fontsize=9)
plt.tight_layout()
plt.savefig('figures/hinh1.png', dpi=150, bbox_inches='tight')
plt.close()

# Figure 2: line chart - growth
fig, ax = plt.subplots(figsize=(6, 3.5))
years = [2020, 2021, 2022, 2023, 2024, 2025]
revenue = [120, 180, 260, 380, 540, 760]
ax.plot(years, revenue, marker='o', linewidth=2, color='#c0504d')
ax.set_xlabel('Năm')
ax.set_ylabel('Doanh thu (triệu USD)')
ax.set_title('Tăng trưởng thị trường AI tại Việt Nam giai đoạn 2020-2025')
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('figures/hinh2.png', dpi=150, bbox_inches='tight')
plt.close()

print("OK")
