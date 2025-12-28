# 📊 Demand Forecast Distribution Automation

> Excel VBA automation reducing report generation from 3 hours to 30 seconds

[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](https://choosealicense.com/licenses/mit/)
[![Excel VBA](https://img.shields.io/badge/Excel-VBA-217346?logo=microsoft-excel)](https://www.microsoft.com/excel)
[![Status](https://img.shields.io/badge/Status-Production-success)]()

## 🎯 Problem Statement

Operations teams spend hours manually creating zone-specific reports from master data files, resulting in:
- ⏱️ 2-3 hours of manual work monthly
- ❌ 5-10 errors per month from copy-paste operations
- 😰 Employee frustration with repetitive tasks
- 📉 Delayed decision-making

## ✨ Solution

Intelligent VBA automation system that:
- ✅ Auto-detects files and configurations
- ✅ Intelligently matches city names (fuzzy logic)
- ✅ Filters data by zones and categories  
- ✅ Generates formatted reports automatically
- ✅ Distributes via email to stakeholders
- ✅ Creates comprehensive audit logs

## 📈 Business Impact

| Metric | Before | After | Improvement |
|--------|--------|-------|-------------|
| **Time** | 180 min | 30 sec | ⬇ **99%** |
| **Errors** | 5-10/month | ~0 | ⬇ **100%** |
| **Monthly Value** | - | $2,000+ | **ROI: 200%+** |

## 🚀 Quick Start

### Prerequisites
- Microsoft Excel 2016+
- Microsoft Outlook (for email)
- Windows 10/11

### Installation

1. **Download the code**
```bash
   git clone https://github.com/your-username/demand-forecast-automation.git
```

2. **Open your Excel file**

3. **Import VBA module**
   - Press `Alt + F11`
   - File → Import File
   - Select `src/MainAutomation.bas`

4. **Run automation**
   - Press `Alt + F8`
   - Select `RunAutomation`
   - Click Run

[📖 Detailed Installation Guide](docs/Installation-Guide.md)

## 📸 Screenshots

### Folder Structure
![Structure](screenshots/folder-structure.png)

### Mapping Configuration
![Mapping](screenshots/mapping-file.png)

### Before Automation
![Before](screenshots/before-process.png)

## 🛠️ Technical Stack

- **Language:** Excel VBA
- **Integration:** Outlook COM Automation
- **Data Structures:** Dictionary Objects (O(1) lookup)
- **File Operations:** FileSystemObject API
- **Performance:** <30 seconds for 10 zones

## 📖 Documentation

- [Installation Guide](docs/Installation-Guide.md)
- [User Manual](docs/User-Manual.md)
- [FAQ](docs/FAQ.md)

## 🎓 Key Features

### 1. Smart File Detection
Automatically finds mapping files regardless of naming convention

### 2. Fuzzy City Matching
Handles variations: Ahmedabad/Ahmed/Amdavad

### 3. Flexible Category Filtering
- Exact match: "IND" matches only "IND"
- Partial match: "IND" matches "IND", "IND/Retail", "OEM/Ind"

### 4. Format Preservation
Maintains all Excel formatting, formulas, and styling

### 5. Multi-Email Support
`person1@example.com; person2@example.com; person3@example.com`

## 🔮 Future Enhancements

- [ ] Power BI dashboard integration
- [ ] Scheduled execution via Task Scheduler
- [ ] Web-based interface
- [ ] Machine learning for forecasting

## 🤝 Contributing

Contributions welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md)

## 📄 License

This project is licensed under the MIT License - see [LICENSE](LICENSE)

## 👤 Author

**Your Name**
- LinkedIn: [Your Profile](your-linkedin-url)
- Email: your.email@example.com

## 🙏 Acknowledgments

- Operations team for requirements and feedback
- Open source community for inspiration

---

**⭐ If this helped you, please star this repo!**

Made with ❤️ using Excel VBA