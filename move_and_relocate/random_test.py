import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Define the colors
colors = {
    "Mint Green": "#98FF98",
    "Teal": "#008080",
    "Turquoise": "#40E0D0"
}

# Create a figure and axis
fig, ax = plt.subplots(figsize=(15, 5))

# Draw rectangles for each color
for i, (color_name, color_hex) in enumerate(colors.items()):
    rect = patches.Rectangle((i * 5, 0), 5, 5, linewidth=1, edgecolor='none', facecolor=color_hex)
    ax.add_patch(rect)
    # Add text in the middle of each rectangle
    ax.text(i * 5 + 2.5, 2.5, color_name, horizontalalignment='center', verticalalignment='center', fontsize=16, color='white' if color_name != "Mint Green" else 'black')

# Set the limits and remove the axes
ax.set_xlim(0, 15)
ax.set_ylim(0, 5)
ax.axis('off')

# Show the plot
plt.show()
