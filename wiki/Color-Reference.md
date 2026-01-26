SSMS EnvTabs supports 16 distinct colors for tab grouping. You can assign these colors to different servers or databases using the `colorIndex` property in your configuration.

![SSMS EnvTabs Colors](./images/colors.png)

## Available Colors

| colorIndex | Color Name | Hex Code | RGB Values |
| :---: | :--- | :--- | :--- |
| **0** | Lavender | `#9083ef` | 144, 131, 239 |
| **1** | Gold | `#d0b132` | 208, 177, 50 |
| **2** | Cyan | `#30b1cd` | 48, 177, 205 |
| **3** | Burgundy | `#cf6468` | 207, 100, 104 |
| **4** | Green | `#6ba12a` | 107, 161, 42 |
| **5** | Brown | `#bc8f6f` | 188, 143, 111 |
| **6** | Royal Blue | `#5bb2fa` | 91, 178, 250 |
| **7** | Pumpkin | `#d67441` | 214, 116, 65 |
| **8** | Gray | `#bdbcbc` | 189, 189, 189 |
| **9** | Volt | `#cbcc38` | 203, 204, 56 |
| **10** | Teal | `#2aa0a4` | 42, 160, 164 |
| **11** | Magenta | `#d957a7` | 217, 87, 167 |
| **12** | Mint | `#6bc6a5` | 107, 198, 165 |
| **13** | Dark Brown | `#946a5b` | 148, 106, 91 |
| **14** | Blue | `#6a8ec6` | 106, 142, 198 |
| **15** | Pink | `#e0a3a5` | 224, 163, 165 |

## Usage Tips

- **-1**: Use `-1` for "Random" color assignment.
- **Environment Grouping**: Consider assigning specific colors to environment types (e.g., Prod, QA, Dev) to quickly identify connection context.
- **Accessibility**: When selecting colors, consider contrast against your SSMS theme (Light or Dark) and potential color vision deficiencies.
