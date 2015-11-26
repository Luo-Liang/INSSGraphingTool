# INSS Graphing Tool
intelliSys Graphing 

1. Hist2CDF.cs

Converts raw data into aggregated CDF.

Given a directory with raw data,

![Directory](https://raw.github.com/Luo-Liang/INSSGraphingTool/master/Figures/Layout.png)

Each file contains many lines. We'd like to extract one sample per line from each file.

![File Structure](https://raw.github.com/Luo-Liang/INSSGraphingTool/master/Figures/RawData.png)

Hist2CDF.exe aggregates CDF by each file, then creates a chart with different CDFs from all files.

![Result](https://raw.github.com/Luo-Liang/INSSGraphingTool/master/Figures/Result.png)

Hist2CDF.exe **Directory** **Stepping** **Minimum** **GroupingRule**

**GroupingRule** allows partition all files into several groups, and one graph is plotted for each group.

You can down sampling by setting stepping to a larger value.
