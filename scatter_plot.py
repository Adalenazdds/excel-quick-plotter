import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.gridspec import GridSpec
from matplotlib.colors import Normalize
import matplotlib.cm as cm

# 1. 模拟生成随机数据
np.random.seed(42)
n_samples = 2000

# 第一组数据 (对应图中的蓝色系)
x1 = np.random.normal(loc=47, scale=2.5, size=n_samples)
y1 = np.random.normal(loc=46, scale=4.0, size=n_samples)

# 第二组数据 (对应图中的红色系)
x2 = np.random.normal(loc=53, scale=2.5, size=n_samples)
y2 = np.random.normal(loc=54, scale=3.5, size=n_samples)

# 2. 创建画布与 4x4 网格布局
fig = plt.figure(figsize=(10, 8))
# width_ratios 适当调小最后一列的比例，避免右侧直方图过宽
gs = GridSpec(4, 4, width_ratios=[1, 1, 1, 0.6], height_ratios=[0.6, 1, 1, 1], wspace=0.4, hspace=0.2)

ax_top = fig.add_subplot(gs[0, 0:3])
ax_main = fig.add_subplot(gs[1:4, 0:3])
ax_right = fig.add_subplot(gs[1:4, 3], sharey=ax_main)

# 3. 绘制中间的 2D KDE 云图
# levels 控制等高线的层数，alpha 控制透明度以便重叠部分可见
sns.kdeplot(x=x1, y=y1, ax=ax_main, cmap="Blues", fill=True, thresh=0.05, alpha=0.7, levels=12)
sns.kdeplot(x=x2, y=y2, ax=ax_main, cmap="Reds", fill=True, thresh=0.05, alpha=0.6, levels=12)

ax_main.set_xlim(30, 70)
ax_main.set_ylim(30, 70)
ax_main.set_xlabel('X')
ax_main.set_ylabel('Y')

# 4. 绘制顶部直方图 (X轴边缘分布)
# 使用 stat="probability" 保持 y 轴在 0~0.15 这个量级
sns.histplot(x=x1, ax=ax_top, color="#A8D1E7", binwidth=0.8, stat="probability", alpha=0.6, edgecolor="#6D9DBB", label="X")
sns.histplot(x=x2, ax=ax_top, color="#F19C9C", binwidth=0.8, stat="probability", alpha=0.6, edgecolor="#C26A6A", label="Y")
ax_top.set_xlim(ax_main.get_xlim())
ax_top.tick_params(labelbottom=False) # 隐藏 X 轴刻度标签
ax_top.set_xlabel('')
ax_top.set_ylabel('Frequency')
ax_top.legend(loc='upper right', frameon=False)
# 去除顶部和右侧的边框使其更清爽
ax_top.spines['top'].set_visible(False)
ax_top.spines['right'].set_visible(False)

# 5. 绘制右侧直方图 (Y轴边缘分布)
sns.histplot(y=y1, ax=ax_right, color="#A8D1E7", binwidth=0.8, stat="probability", alpha=0.6, edgecolor="#6D9DBB")
sns.histplot(y=y2, ax=ax_right, color="#F19C9C", binwidth=0.8, stat="probability", alpha=0.6, edgecolor="#C26A6A")
ax_right.tick_params(labelleft=False) # 隐藏 Y 轴刻度标签
ax_right.set_ylabel('')
ax_right.set_xlabel('Frequency')
ax_right.spines['top'].set_visible(False)
ax_right.spines['right'].set_visible(False)

# 6. 精细调整单独构建 Colorbar 的位置
# 强制渲染一次画布，以便准确获取 ax_main 的实际像素位置
fig.canvas.draw()
pos = ax_main.get_position()

# 设定 Colorbar 的宽高与位置 (使用相对坐标)
cb_width = 0.015
cb_height = pos.height * 0.4
cb_x = pos.x1 + 0.02 # 紧靠主图右侧边缘

# 顶部蓝色的 Colorbar
cax_blue = fig.add_axes([cb_x, pos.y0 + pos.height * 0.55, cb_width, cb_height])
sm_blue = cm.ScalarMappable(cmap="Blues", norm=Normalize(vmin=0, vmax=1.0))
sm_blue.set_array([])
fig.colorbar(sm_blue, cax=cax_blue)

# 底部红色的 Colorbar
cax_red = fig.add_axes([cb_x, pos.y0 + pos.height * 0.05, cb_width, cb_height])
sm_red = cm.ScalarMappable(cmap="Reds", norm=Normalize(vmin=0, vmax=1.0))
sm_red.set_array([])
fig.colorbar(sm_red, cax=cax_red)

# 确保布局不被破坏，这里不使用 tight_layout，因为我们用 add_axes 强行插入了组件
plt.show()