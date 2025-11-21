---
title: 'Charts: Create and customize Excel charts with Office Scripts'
description: A collection of code samples showing how to create and customize different types of Excel charts in Office Scripts.
ms.date: 11/20/2025
ms.localizationpriority: medium
---

# Charts: Create and customize Excel charts with Office Scripts

This article demonstrates how to create every type of Excel chart supported by Office Scripts. Each sample includes data and showcases APIs specific to that chart type. Use these samples as starting points for your own charting solutions.

> [!TIP]
> Run each sample on an empty worksheet for the best results.

## Column charts

Column charts display data as vertical bars, making them ideal for comparing values across categories.

### Column clustered chart

This sample creates a clustered column chart comparing quarterly sales across different products.

:::image type="content" source="../../images/column-clustered-chart.png" alt-text="A clustered column chart showing quarterly sales data for laptops, tablets, and phones across four quarters.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Product", "Q1", "Q2", "Q3", "Q4"],
    ["Laptops", 45000, 52000, 48000, 61000],
    ["Tablets", 32000, 35000, 38000, 42000],
    ["Phones", 28000, 31000, 29000, 35000]
  ];
  const dataRange = sheet.getRange("A1:E4");
  dataRange.setValues(data);
  
  // Create column clustered chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.columnClustered,
    dataRange
  );
  chart.setPosition("A6");
  chart.getTitle().setText("Quarterly Sales by Product");
  
  // Customize the chart.
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
  chart.getAxes().getValueAxis().setDisplayUnit(ExcelScript.ChartAxisDisplayUnit.thousands);
}
```

### Column stacked chart

This sample creates a stacked column chart showing composition of sales by region.

:::image type="content" source="../../images/column-stacked-chart.png" alt-text="A stacked column chart displaying regional sales contributions from North, South, East, and West regions across four months.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Month", "North", "South", "East", "West"],
    ["Jan", 15000, 18000, 22000, 19000],
    ["Feb", 17000, 19000, 21000, 20000],
    ["Mar", 19000, 21000, 23000, 22000],
    ["Apr", 21000, 22000, 25000, 24000]
  ];
  const dataRange = sheet.getRange("A1:E5");
  dataRange.setValues(data);
  
  // Create stacked column chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.columnStacked,
    dataRange
  );
  chart.setPosition("A7");
  chart.getTitle().setText("Regional Sales Contribution");
  
  // Customize chart.
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
  chart.getAxes().getCategoryAxis().setReversePlotOrder(false);
}
```

### Column stacked 100% chart

This sample creates a 100% stacked column chart showing percentage distribution.

:::image type="content" source="../../images/column-stacked-100-chart.png" alt-text="A 100% stacked column chart showing the percentage distribution of desktop, mobile, and tablet usage across four quarters.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Category", "Desktop", "Mobile", "Tablet"],
    ["Q1", 5500, 3200, 1300],
    ["Q2", 4800, 3800, 1400],
    ["Q3", 4200, 4100, 1700],
    ["Q4", 3900, 4500, 1600]
  ];
  const dataRange = sheet.getRange("A1:D5");
  dataRange.setValues(data);
  
  // Create 100% stacked column chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.columnStacked100,
    dataRange
  );
  chart.setPosition("A7");
  chart.getTitle().setText("Device Usage Distribution");
}
```

## Bar charts

Bar charts display data as horizontal bars, useful when category names are long or you have many categories.

### Bar clustered chart

This sample creates a clustered bar chart comparing employee performance ratings.

:::image type="content" source="../../images/bar-clustered-chart.png" alt-text="A clustered bar chart comparing technical, communication, and leadership ratings for five employees.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Employee", "Technical", "Communication", "Leadership"],
    ["Chantha Mean", 9.2, 8.8, 9.5],
    ["Marcio Alvez", 8.5, 9.1, 8.2],
    ["Sobia Khanam", 9.8, 9.3, 9.6],
    ["Altynbek Joldubai", 8.9, 8.5, 8.8],
    ["Adriana Mota", 9.5, 9.7, 9.2]
  ];
  const dataRange = sheet.getRange("A1:D6");
  dataRange.setValues(data);
  
  // Create bar clustered chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.barClustered,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Employee Performance Ratings");
  
  // Customize.
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
  chart.getAxes().getValueAxis().setMinimum(0);
  chart.getAxes().getValueAxis().setMaximum(10);
}
```

### Bar stacked chart

:::image type="content" source="../../images/bar-stacked-chart.png" alt-text="A stacked bar chart showing project timelines with planning, development, testing, and deployment phases for three projects.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Project Phase", "Planning", "Development", "Testing", "Deployment"],
    ["Project A", 2, 6, 3, 1],
    ["Project B", 3, 8, 4, 2],
    ["Project C", 1, 5, 2, 1]
  ];
  const dataRange = sheet.getRange("A1:E4");
  dataRange.setValues(data);
  
  // Create bar stacked chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.barStacked,
    dataRange
  );
  chart.setPosition("A6");
  chart.getTitle().setText("Project Timeline (Weeks)");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

## Line charts

Line charts show trends over time, perfect for displaying continuous data.

### Line chart

This sample creates a basic line chart showing temperature trends.

:::image type="content" source="../../images/line-chart.png" alt-text="A line chart displaying average monthly temperature trends throughout the year.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Month", "Avg Temp (°F)"],
    ["Jan", 42],
    ["Feb", 45],
    ["Mar", 52],
    ["Apr", 61],
    ["May", 70],
    ["Jun", 78],
    ["Jul", 84],
    ["Aug", 82],
    ["Sep", 75],
    ["Oct", 64],
    ["Nov", 53],
    ["Dec", 45]
  ];
  const dataRange = sheet.getRange("A1:B13");
  dataRange.setValues(data);
  
  // Create line chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.line,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Average Monthly Temperature");
  
  // Customize line chart.
  chart.getAxes().getCategoryAxis().setTickLabelPosition(
    ExcelScript.ChartAxisTickLabelPosition.low
  );
  chart.getAxes().getValueAxis().getMajorGridlines().getFormat().getLine().setColor("gray");
}
```

### Line with markers chart

This sample creates a line chart with markers to emphasize individual data points.

:::image type="content" source="../../images/line-markers-chart.png" alt-text="A line chart with circular markers showing weekly visitor growth for three websites over six weeks.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Week", "Website A", "Website B", "Website C"],
    ["Week 1", 1250, 980, 1100],
    ["Week 2", 1380, 1050, 1150],
    ["Week 3", 1520, 1180, 1320],
    ["Week 4", 1690, 1280, 1450],
    ["Week 5", 1850, 1420, 1580],
    ["Week 6", 2020, 1590, 1720]
  ];
  const dataRange = sheet.getRange("A1:D7");
  dataRange.setValues(data);
  
  // Create line chart with markers.
  const chart = sheet.addChart(
    ExcelScript.ChartType.lineMarkers,
    dataRange
  );
  chart.setPosition("A9");
  chart.getTitle().setText("Weekly Visitor Growth");
  
  // Customize markers.
  const series = chart.getSeries();
  series.forEach((s) => {
    s.setMarkerSize(8);
    s.setMarkerStyle(ExcelScript.ChartMarkerStyle.circle);
  });
}
```

### Line stacked chart

:::image type="content" source="../../images/line-stacked-chart.png" alt-text="A stacked line chart showing cumulative product revenue for three products across four quarters.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Quarter", "Product A", "Product B", "Product C"],
    ["Q1 2024", 25000, 18000, 12000],
    ["Q2 2024", 28000, 20000, 14000],
    ["Q3 2024", 32000, 22000, 16000],
    ["Q4 2024", 35000, 25000, 18000]
  ];
  const dataRange = sheet.getRange("A1:D5");
  dataRange.setValues(data);
  
  // Create stacked line chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.lineStacked,
    dataRange
  );
  chart.setPosition("A7");
  chart.getTitle().setText("Cumulative Product Revenue");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.top);
}
```

## Pie charts

Pie charts show proportional relationships in a dataset, displaying each value as a slice of the whole.

### Pie chart

This sample creates a pie chart showing market share distribution.

:::image type="content" source="../../images/pie-chart.png" alt-text="A pie chart displaying market share distribution with percentage labels for five companies.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data with multiple slices.
  const data = [
    ["Company", "Market Share"],
    ["Alpha Corp", 28.5],
    ["Beta Inc", 22.3],
    ["Gamma LLC", 18.7],
    ["Delta Co", 15.2],
    ["Others", 15.3]
  ];
  const dataRange = sheet.getRange("A1:B6");
  dataRange.setValues(data);
  
  // Create pie chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.pie,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Market Share Distribution");
  
  // Customize pie chart.
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
  
  // Add data labels showing percentages.
  const series = chart.getSeries()[0];
  series.setHasDataLabels(true);
  series.getDataLabels().setShowPercentage(true);
  series.getDataLabels().setShowSeriesName(false);
  series.getDataLabels().setShowCategoryName(false);
  series.getDataLabels().setShowValue(false);
}
```

### Exploded pie chart

This sample creates an exploded pie chart with separated slices for emphasis.

:::image type="content" source="../../images/exploded-pie-chart.png" alt-text="An exploded pie chart showing monthly budget breakdown with separated slices for housing, transportation, food, and other expenses.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Expense Category", "Amount"],
    ["Housing", 1800],
    ["Transportation", 650],
    ["Food", 550],
    ["Utilities", 200],
    ["Entertainment", 300],
    ["Savings", 500]
  ];
  const dataRange = sheet.getRange("A1:B7");
  dataRange.setValues(data);
  
  // Create exploded pie chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.pieExploded,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Monthly Budget Breakdown");
  
  // Customize.
  const series = chart.getSeries()[0];
  series.setHasDataLabels(true);
  series.getDataLabels().setShowCategoryName(true);
  series.getDataLabels().setShowValue(true);
  series.getDataLabels().setPosition(ExcelScript.ChartDataLabelPosition.bestFit);
}
```

## Doughnut charts

Doughnut charts are similar to pie charts but can display multiple data series and have a hole in the center.

### Doughnut chart

:::image type="content" source="../../images/doughnut-chart.png" alt-text="A doughnut chart displaying revenue by category with percentage labels for hardware, software, services, training, and support.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Category", "2024"],
    ["Hardware", 45000],
    ["Software", 32000],
    ["Services", 28000],
    ["Training", 15000],
    ["Support", 20000]
  ];
  const dataRange = sheet.getRange("A1:B6");
  dataRange.setValues(data);
  
  // Create doughnut chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.doughnut,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Revenue by Category");
  
  // Customize doughnut chart.
  const series = chart.getSeries()[0];
  series.setHasDataLabels(true);
  series.getDataLabels().setShowPercentage(true);
  series.getDataLabels().setShowLeaderLines(true);
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

### Exploded doughnut chart

:::image type="content" source="../../images/exploded-doughnut-chart.png" alt-text="An exploded doughnut chart showing regional sales distribution with separated segments for different geographical regions.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Region", "Sales %"],
    ["North America", 35.2],
    ["Europe", 28.7],
    ["Asia Pacific", 22.3],
    ["Latin America", 8.5],
    ["Others", 5.3]
  ];
  const dataRange = sheet.getRange("A1:B6");
  dataRange.setValues(data);
  
  // Create exploded doughnut chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.doughnutExploded,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Regional Sales Distribution");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
}
```

## Area charts

Area charts emphasize the magnitude of change over time and show the cumulative total.

### Area chart

:::image type="content" source="../../images/area-chart.png" alt-text="An area chart showing monthly revenue trends with a filled area under the line.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Month", "Revenue"],
    ["Jan", 45000],
    ["Feb", 52000],
    ["Mar", 48000],
    ["Apr", 61000],
    ["May", 58000],
    ["Jun", 67000]
  ];
  const dataRange = sheet.getRange("A1:B7");
  dataRange.setValues(data);
  
  // Create area chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.area,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Monthly Revenue Trend");
  
  // Customize area chart.
  chart.getAxes().getValueAxis().setDisplayUnit(
    ExcelScript.ChartAxisDisplayUnit.thousands
  );
}
```

### Stacked area chart

:::image type="content" source="../../images/stacked-area-chart.png" alt-text="A stacked area chart displaying renewable energy production by source (solar, wind, hydro, geothermal) over five years.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["", "Solar", "Wind", "Hydro", "Geothermal"],
    ["2020", 150, 120, 200, 30],
    ["2021", 180, 145, 205, 35],
    ["2022", 220, 175, 210, 40],
    ["2023", 270, 210, 215, 48],
    ["2024", 330, 250, 220, 55]
  ];
  const dataRange = sheet.getRange("A1:E6");
  dataRange.setValues(data);
  
  // Create stacked area chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.areaStacked,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Renewable Energy Production (TWh)");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.top);
}
```

## Scatter (XY) charts

Scatter charts show relationships between two numerical variables, ideal for correlation analysis.

### Scatter chart

:::image type="content" source="../../images/scatter-chart.png" alt-text="A scatter chart showing the relationship between study hours and test scores.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Study Hours", "Test Score"],
    [2, 65],
    [3, 68],
    [4, 72],
    [5, 75],
    [5.5, 78],
    [6, 82],
    [7, 85],
    [7.5, 88],
    [8, 90],
    [9, 92],
    [10, 95]
  ];
  const dataRange = sheet.getRange("A1:B12");
  dataRange.setValues(data);
  
  // Create scatter chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.xyscatter,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Study Hours vs Test Scores");
  
  // Customize scatter chart.
  chart.getAxes().getCategoryAxis().setDisplayUnit(ExcelScript.ChartAxisDisplayUnit.none);
  chart.getAxes().getValueAxis().setMinimum(0);
  chart.getAxes().getValueAxis().setMaximum(100);
  
  // Remove legend as there's only one series.
  chart.getLegend().setVisible(false);
}
```

### Scatter with lines chart

:::image type="content" source="../../images/scatter-lines-chart.png" alt-text="A scatter chart with connecting lines showing ice cream sales versus temperature.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Temperature (°C)", "Sales ($)"],
    [15, 2500],
    [18, 2800],
    [22, 3200],
    [25, 3800],
    [28, 4500],
    [32, 5200],
    [35, 5800],
    [38, 6100]
  ];
  const dataRange = sheet.getRange("A1:B9");
  dataRange.setValues(data);
  
  // Create scatter chart with lines.
  const chart = sheet.addChart(
    ExcelScript.ChartType.xyscatterLines,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Ice Cream Sales vs Temperature");
  
  chart.getLegend().setVisible(false);
}
```

### Scatter with smooth lines chart

:::image type="content" source="../../images/scatter-smooth-lines-chart.png" alt-text="A scatter chart with smooth curved lines comparing growth patterns of two series.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["X", "Series 1", "Series 2"],
    [1, 10, 15],
    [2, 15, 18],
    [3, 25, 22],
    [4, 40, 28],
    [5, 60, 35],
    [6, 85, 43],
    [7, 115, 52]
  ];
  const dataRange = sheet.getRange("A1:C8");
  dataRange.setValues(data);
  
  // Create scatter chart with smooth lines.
  const chart = sheet.addChart(
    ExcelScript.ChartType.xyscatterSmooth,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Growth Comparison");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

## Bubble charts

Bubble charts display three dimensions of data: X and Y position plus bubble size.

### Bubble chart

:::image type="content" source="../../images/bubble-chart.png" alt-text="A bubble chart showing product analysis with price, quality score, and market share represented by bubble size.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data for bubble chart.
  // Each product is a separate data series with X, Y, and Size values.
  const data = [
    ["Product", "Price ($)", "Quality Score", "Market Share"],
    ["Laptops", 150, 85, 25],
    ["Tablets", 200, 90, 40],
    ["Phones", 100, 75, 15],
    ["Monitors", 250, 92, 35],
    ["Keyboards", 175, 88, 30],
    ["Mice", 120, 80, 20]
  ];
  const dataRange = sheet.getRange("A1:D7");
  dataRange.setValues(data);
  
  // Create bubble chart - manually add each series.
  const chart = sheet.addChart(
    ExcelScript.ChartType.bubble,
    sheet.getRange("B1:D1") // Start with just headers to create empty chart.
  );
  chart.setPosition("A9");
  chart.getTitle().setText("Product Analysis: Price vs Quality (Market Share)");
  
  // Remove any default series that were created.
  while (chart.getSeries().length > 0) {
    chart.getSeries()[0].delete();
  }
  
  // Add each product as its own series.
  for (let i = 2; i <= 7; i++) {
    const productName = sheet.getRange(`A${ i } `).getValue() as string;
    const newSeries = chart.addChartSeries();
    
    newSeries.setName(productName);
    newSeries.setXAxisValues(sheet.getRange(`B${ i }:B${ i } `));
    newSeries.setValues(sheet.getRange(`C${ i }:C${ i } `));
    newSeries.setBubbleSizes(sheet.getRange(`D${ i }:D${ i } `));
  }
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
}
```

### 3D bubble chart

:::image type="content" source="../../images/3d-bubble-chart.png" alt-text="A 3D bubble chart comparing cities by population, GDP per capita, and area represented by bubble size.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data for 3D bubble chart.
  // Each city is a separate data series with X, Y, and Size values.
  const data = [
    ["City", "Population (millions)", "GDP per Capita ($k)", "Area (km²)"],
    ["Tokyo", 8.5, 65, 300],
    ["Berlin", 3.2, 58, 150],
    ["Sydney", 5.8, 72, 250],
    ["Toronto", 2.1, 52, 100],
    ["Singapore", 4.5, 68, 180]
  ];
  const dataRange = sheet.getRange("A1:D6");
  dataRange.setValues(data);
  
  // Create 3D bubble chart - manually add each series.
  const chart = sheet.addChart(
    ExcelScript.ChartType.bubble3DEffect,
    sheet.getRange("B1:D1") // Start with just headers to create empty chart.
  );
  chart.setPosition("A8");
  chart.getTitle().setText("City Comparison: Population vs GDP per Capita (Area)");
  
  // Remove any default series that were created.
  while (chart.getSeries().length > 0) {
    chart.getSeries()[0].delete();
  }
  
  // Add each city as its own series.
  for (let i = 2; i <= 6; i++) {
    const cityName = sheet.getRange(`A${ i } `).getValue() as string;
    const newSeries = chart.addChartSeries();
    
    newSeries.setName(cityName);
    newSeries.setXAxisValues(sheet.getRange(`B${ i }:B${ i } `));
    newSeries.setValues(sheet.getRange(`C${ i }:C${ i } `));
    newSeries.setBubbleSizes(sheet.getRange(`D${ i }:D${ i } `));
  }
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

## Stock charts

Stock charts display financial data with high, low, close (and optionally open and volume) values.

### Stock HLC chart

This sample creates a High-Low-Close stock chart.

:::image type="content" source="../../images/stock-hlc-chart.png" alt-text="A High-Low-Close stock chart showing stock price movements with high, low, and close values for five trading days.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data (Date, High, Low, Close).
  const data = [
    ["Date", "High", "Low", "Close"],
    ["1/1/2024", 152.50, 148.20, 151.00],
    ["1/2/2024", 153.80, 150.50, 152.30],
    ["1/3/2024", 154.20, 151.80, 153.50],
    ["1/4/2024", 155.00, 152.30, 153.80],
    ["1/5/2024", 156.50, 153.50, 155.20]
  ];
  const dataRange = sheet.getRange("A1:D6");
  dataRange.setValues(data);
  
  // Create HLC stock chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.stockHLC,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Stock Price Movement (HLC)");
  
  // Customize stock chart.
  chart.getAxes().getValueAxis().setMinimum(145);
  chart.getAxes().getValueAxis().setMaximum(160);
}
```

### Stock OHLC chart

This sample creates an Open-High-Low-Close stock chart.

:::image type="content" source="../../images/stock-ohlc-chart.png" alt-text="An Open-High-Low-Close stock chart displaying stock price movements with open, high, low, and close values.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data (Date, Open, High, Low, Close).
  const data = [
    ["Date", "Open", "High", "Low", "Close"],
    ["1/1/2024", 150.00, 152.50, 148.20, 151.00],
    ["1/2/2024", 151.00, 153.80, 150.50, 152.30],
    ["1/3/2024", 152.30, 154.20, 151.80, 153.50],
    ["1/4/2024", 153.50, 155.00, 152.30, 153.80],
    ["1/5/2024", 153.80, 156.50, 153.50, 155.20]
  ];
  const dataRange = sheet.getRange("A1:E6");
  dataRange.setValues(data);
  
  // Create OHLC stock chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.stockOHLC,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Stock Price Movement (OHLC)");
  
  chart.getAxes().getValueAxis().setMinimum(145);
  chart.getAxes().getValueAxis().setMaximum(160);
}
```

### Stock VHLC chart

This sample creates a Volume-High-Low-Close stock chart.

:::image type="content" source="../../images/stock-vhlc-chart.png" alt-text="A Volume-High-Low-Close stock chart showing trading volume alongside price movements.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data (Date, Volume, High, Low, Close).
  const data = [
    ["Date", "Volume", "High", "Low", "Close"],
    ["1/1/2024", 2500000, 152.50, 148.20, 151.00],
    ["1/2/2024", 3200000, 153.80, 150.50, 152.30],
    ["1/3/2024", 2800000, 154.20, 151.80, 153.50],
    ["1/4/2024", 3500000, 155.00, 152.30, 153.80],
    ["1/5/2024", 4100000, 156.50, 153.50, 155.20]
  ];
  const dataRange = sheet.getRange("A1:E6");
  dataRange.setValues(data);
  
  // Create VHLC stock chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.stockVHLC,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Stock Price with Volume (VHLC)");
  
  // Customize to improve visibility of stock lines against volume bars.
  chart.getAxes().getValueAxis().setMinimum(145);
  chart.getAxes().getValueAxis().setMaximum(160);
  
  // Make the volume bars more transparent or lighter colored.
  const volumeSeries = chart.getSeries()[0];
  volumeSeries.getFormat().getFill().setSolidColor("#B0C4DE"); // Light steel blue.
}
```

### Stock VOHLC chart

This sample creates a Volume-Open-High-Low-Close stock chart.

:::image type="content" source="../../images/stock-vohlc-chart.png" alt-text="A Volume-Open-High-Low-Close stock chart displaying complete stock analysis with volume and all price points.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data (Date, Volume, Open, High, Low, Close).
  const data = [
    ["Date", "Volume", "Open", "High", "Low", "Close"],
    ["1/1/2024", 2500000, 150.00, 152.50, 148.20, 151.00],
    ["1/2/2024", 3200000, 151.00, 153.80, 150.50, 152.30],
    ["1/3/2024", 2800000, 152.30, 154.20, 151.80, 153.50],
    ["1/4/2024", 3500000, 153.50, 155.00, 152.30, 153.80],
    ["1/5/2024", 4100000, 153.80, 156.50, 153.50, 155.20]
  ];
  const dataRange = sheet.getRange("A1:F6");
  dataRange.setValues(data);
  
  // Create VOHLC stock chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.stockVOHLC,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Complete Stock Analysis (VOHLC)");
  
  // Customize to improve visibility of stock lines against volume bars.
  chart.getAxes().getValueAxis().setMinimum(145);
  chart.getAxes().getValueAxis().setMaximum(160);
  
  // Make the volume bars more transparent or lighter colored.
  const volumeSeries = chart.getSeries()[0];
  volumeSeries.getFormat().getFill().setSolidColor("#B0C4DE"); // Light steel blue.
}
```

## Radar charts

Radar charts display multivariate data on axes starting from the same point, useful for comparing multiple variables.

### Radar chart

:::image type="content" source="../../images/radar-chart.png" alt-text="A radar chart comparing multiple attributes of two products across six dimensions.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Attribute", "Product A", "Product B"],
    ["Speed", 8, 6],
    ["Reliability", 9, 8],
    ["Cost", 6, 9],
    ["Features", 7, 8],
    ["Support", 8, 7],
    ["Ease of Use", 9, 9]
  ];
  const dataRange = sheet.getRange("A1:C7");
  dataRange.setValues(data);
  
  // Create radar chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.radar,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Product Comparison Radar");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

### Radar with markers chart

:::image type="content" source="../../images/radar-markers-chart.png" alt-text="A radar chart with markers showing developer skill assessments across different technologies.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Skill", "Danaite Alemseged", "Fazekas Peter", "Jeffry Goh"],
    ["JavaScript", 9, 7, 8],
    ["Python", 7, 9, 6],
    ["SQL", 8, 8, 9],
    ["Cloud", 8, 6, 7],
    ["DevOps", 6, 8, 9]
  ];
  const dataRange = sheet.getRange("A1:D6");
  dataRange.setValues(data);
  
  // Create radar chart with markers.
  const chart = sheet.addChart(
    ExcelScript.ChartType.radarMarkers,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Developer Skill Assessment");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
}
```

### Filled radar chart

:::image type="content" source="../../images/filled-radar-chart.png" alt-text="A filled radar chart comparing current performance versus target metrics across six categories.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data - Target first, then Current so Current is drawn on top.
  const data = [
    ["Category", "Target", "Current"],
    ["Customer Satisfaction", 9.0, 7.5],
    ["Product Quality", 9.5, 8.2],
    ["Delivery Speed", 9.0, 6.8],
    ["Price Competitiveness", 8.5, 7.0],
    ["Innovation", 9.5, 8.5],
    ["Market Presence", 9.0, 7.2]
  ];
  const dataRange = sheet.getRange("A1:C7");
  dataRange.setValues(data);
  
  // Create filled radar chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.radarFilled,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Performance vs Target Metrics");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

## Treemap chart

Treemap charts display hierarchical data as nested rectangles, with size representing values.

:::image type="content" source="../../images/treemap-chart.png" alt-text="A treemap chart showing sales by category and subcategory with nested rectangles sized by value.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Category", "Subcategory", "Value"],
    ["Electronics", "Phones", 45000],
    ["Electronics", "Laptops", 38000],
    ["Electronics", "Tablets", 22000],
    ["Furniture", "Desks", 18000],
    ["Furniture", "Chairs", 25000],
    ["Furniture", "Storage", 12000],
    ["Clothing", "Shirts", 15000],
    ["Clothing", "Pants", 18000],
    ["Clothing", "Accessories", 8000]
  ];
  const dataRange = sheet.getRange("A1:C10");
  dataRange.setValues(data);
  
  // Create treemap chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.treemap,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Sales by Category (Treemap)");
  
  // Treemap-specific customization.
  chart.getLegend().setVisible(false);
}
```

## Sunburst chart

Sunburst charts display hierarchical data in concentric circles, with each level represented by a ring.

:::image type="content" source="../../images/sunburst-chart.png" alt-text="A sunburst chart displaying organizational revenue breakdown with hierarchical levels shown in concentric rings.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Level 1", "Level 2", "Level 3", "Value"],
    ["Sales", "North America", "USA", 125000],
    ["Sales", "North America", "Canada", 35000],
    ["Sales", "Europe", "UK", 45000],
    ["Sales", "Europe", "Germany", 52000],
    ["Sales", "Europe", "France", 38000],
    ["Marketing", "Digital", "Social Media", 28000],
    ["Marketing", "Digital", "Email", 15000],
    ["Marketing", "Traditional", "Print", 12000],
    ["Marketing", "Traditional", "TV", 35000]
  ];
  const dataRange = sheet.getRange("A1:D10");
  dataRange.setValues(data);
  
  // Create sunburst chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.sunburst,
    dataRange
  );
  chart.setPosition("E1");
  chart.getTitle().setText("Organizational Revenue Breakdown");
  
  chart.getLegend().setVisible(false);
}
```

## Waterfall chart

Waterfall charts show how an initial value is affected by positive and negative values, displaying the cumulative effect.

:::image type="content" source="../../images/waterfall-chart.png" alt-text="A waterfall chart showing profit and loss analysis with starting balance, revenue, expenses, and ending balance.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Category", "Amount"],
    ["Starting Balance", 50000],
    ["Revenue", 125000],
    ["Cost of Goods", -45000],
    ["Operating Expenses", -32000],
    ["Marketing", -15000],
    ["Taxes", -18000],
    ["Ending Balance", 65000]
  ];
  const dataRange = sheet.getRange("A1:B8");
  dataRange.setValues(data);
  
  // Create waterfall chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.waterfall,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Profit & Loss Analysis");
  
  // Waterfall charts automatically handle the flow visualization.
  chart.getLegend().setVisible(false);
}
```

## Funnel chart

Funnel charts show progressive reduction of data through stages, commonly used in sales and conversion analysis.

:::image type="content" source="../../images/funnel-chart.png" alt-text="A funnel chart displaying sales conversion stages from website visitors to completed purchases.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Stage", "Count"],
    ["Website Visitors", 10000],
    ["Product Page Views", 4500],
    ["Add to Cart", 1200],
    ["Started Checkout", 800],
    ["Completed Purchase", 450]
  ];
  const dataRange = sheet.getRange("A1:B6");
  dataRange.setValues(data);
  
  // Create funnel chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.funnel,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Sales Funnel Conversion");
}
```

## Box and whisker chart

Box and whisker charts show distribution of data through quartiles, highlighting median and outliers.

:::image type="content" source="../../images/box-whisker-chart.png" alt-text="A box and whisker chart showing response time distribution by region with quartiles and outliers.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data (multiple observations per category).
  const data = [
    ["Region", "Response Time (ms)"],
    ["North", 245],
    ["North", 268],
    ["North", 232],
    ["North", 289],
    ["North", 251],
    ["South", 312],
    ["South", 298],
    ["South", 334],
    ["South", 305],
    ["South", 321],
    ["East", 198],
    ["East", 215],
    ["East", 187],
    ["East", 223],
    ["East", 206],
    ["West", 267],
    ["West", 281],
    ["West", 254],
    ["West", 273],
    ["West", 269]
  ];
  const dataRange = sheet.getRange("A1:B21");
  dataRange.setValues(data);
  
  // Create box and whisker chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.boxwhisker,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Response Time Distribution by Region");
  
  chart.getLegend().setVisible(false);
}
```

## Histogram chart

Histogram charts display the distribution of numerical data by grouping values into bins.

:::image type="content" source="../../images/histogram-chart.png" alt-text="A histogram chart showing customer age distribution with automatically determined bin ranges.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data (customer ages).
  const data = [
    ["Customer", "Age"],
    ["Customer 1", 23],
    ["Customer 2", 45],
    ["Customer 3", 31],
    ["Customer 4", 52],
    ["Customer 5", 28],
    ["Customer 6", 67],
    ["Customer 7", 38],
    ["Customer 8", 41],
    ["Customer 9", 29],
    ["Customer 10", 55],
    ["Customer 11", 33],
    ["Customer 12", 48],
    ["Customer 13", 26],
    ["Customer 14", 62],
    ["Customer 15", 35],
    ["Customer 16", 44],
    ["Customer 17", 58],
    ["Customer 18", 37],
    ["Customer 19", 49],
    ["Customer 20", 71],
    ["Customer 21", 24],
    ["Customer 22", 39],
    ["Customer 23", 56],
    ["Customer 24", 42],
    ["Customer 25", 64]
  ];
  const dataRange = sheet.getRange("A1:B26");
  dataRange.setValues(data);
  
  // Create histogram chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.histogram,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Customer Age Distribution");
  chart.getLegend().setVisible(false);
}
```

## Pareto chart

Pareto charts combine column and line charts to show both individual values and cumulative totals, following the 80/20 principle.

:::image type="content" source="../../images/pareto-chart.png" alt-text="A pareto chart analyzing defect types with columns showing frequency and a line showing cumulative percentage.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Defect Type", "Frequency"],
    ["Packaging", 45],
    ["Assembly", 32],
    ["Quality Control", 28],
    ["Material", 18],
    ["Design", 12],
    ["Shipping", 8],
    ["Documentation", 5],
    ["Other", 3]
  ];
  const dataRange = sheet.getRange("A1:B9");
  dataRange.setValues(data);
  
  // Create pareto chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.pareto,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Defect Analysis (Pareto)");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

## Surface charts

Surface charts display trends in values across two dimensions in a continuous curve, useful for finding optimal combinations.

### Surface chart

:::image type="content" source="../../images/surface-chart.png" alt-text="A 3D surface chart displaying temperature distribution across two-dimensional coordinates.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data (temperature at different coordinates).
  const data = [
    ["", "0", "5", "10", "15", "20"],
    ["0", 20, 22, 25, 28, 30],
    ["5", 22, 24, 27, 30, 32],
    ["10", 25, 27, 30, 33, 35],
    ["15", 28, 30, 33, 36, 38],
    ["20", 30, 32, 35, 38, 40]
  ];
  const dataRange = sheet.getRange("A1:F6");
  dataRange.setValues(data);
  
  // Create surface chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.surface,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Temperature Distribution (3D Surface)");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
}
```

### Surface wireframe chart

:::image type="content" source="../../images/surface-wireframe-chart.png" alt-text="A wireframe surface chart showing data surface with grid lines and no fill.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["", "A", "B", "C", "D", "E"],
    ["1", 10, 15, 20, 15, 10],
    ["2", 15, 25, 35, 25, 15],
    ["3", 20, 35, 50, 35, 20],
    ["4", 15, 25, 35, 25, 15],
    ["5", 10, 15, 20, 15, 10]
  ];
  const dataRange = sheet.getRange("A1:F6");
  dataRange.setValues(data);
  
  // Create surface wireframe chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.surfaceWireframe,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Data Surface (Wireframe)");
  
  chart.getLegend().setVisible(false);
}
```

### Contour chart (Surface top view)

:::image type="content" source="../../images/contour-chart.png" alt-text="A contour chart displaying signal strength as a top-down view with color-coded regions.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["", "0°", "45°", "90°", "135°", "180°"],
    ["0m", 100, 95, 85, 92, 98],
    ["10m", 95, 88, 75, 85, 93],
    ["20m", 85, 78, 62, 75, 83],
    ["30m", 92, 85, 75, 82, 90],
    ["40m", 98, 93, 83, 90, 96]
  ];
  const dataRange = sheet.getRange("A1:F6");
  dataRange.setValues(data);
  
  // Create contour (top view) chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.surfaceTopView,
    dataRange
  );
  chart.setPosition("A8");
  chart.getTitle().setText("Signal Strength Map (Contour)");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
}
```

## Region map chart

Region map charts (also called filled map charts) display values across geographical regions.

:::image type="content" source="../../images/region-map-chart.png" alt-text="A region map chart showing sales by state with color-coded geographical regions.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data (state names and values).
  const data = [
    ["State", "Sales"],
    ["California", 450000],
    ["Texas", 380000],
    ["Florida", 320000],
    ["New York", 410000],
    ["Illinois", 280000],
    ["Pennsylvania", 240000],
    ["Ohio", 220000],
    ["Georgia", 195000],
    ["North Carolina", 185000],
    ["Michigan", 175000]
  ];
  const dataRange = sheet.getRange("A1:B11");
  dataRange.setValues(data);
  
  // Create region map chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.regionMap,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Sales by State");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
}
```

## Combo charts (Bar/Pie combination)

### Bar of pie chart

Bar of pie charts break out smaller slices from a pie chart into a separate bar for better visibility.

:::image type="content" source="../../images/bar-pie-chart.png" alt-text="A bar of pie chart showing category distribution with smaller slices expanded into a separate bar.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Category", "Amount"],
    ["Category A", 45],
    ["Category B", 30],
    ["Category C", 15],
    ["Category D", 4],
    ["Category E", 3],
    ["Category F", 2],
    ["Category G", 1]
  ];
  const dataRange = sheet.getRange("A1:B8");
  dataRange.setValues(data);
  
  // Create bar of pie chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.barOfPie,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Category Distribution (Bar of Pie)");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

### Pie of pie chart

Pie of pie charts display a secondary pie chart showing details of smaller slices.

:::image type="content" source="../../images/pie-pie-chart.png" alt-text="A pie of pie chart showing market segment analysis with a secondary pie expanding smaller segments.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Segment", "Value"],
    ["Enterprise", 425000],
    ["Small Business", 285000],
    ["Consumer", 190000],
    ["Education", 65000],
    ["Government", 55000],
    ["Non-Profit", 30000]
  ];
  const dataRange = sheet.getRange("A1:B7");
  dataRange.setValues(data);
  
  // Create pie of pie chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.pieOfPie,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Market Segment Analysis (Pie of Pie)");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
}
```

## 3D chart variations

### Cone column chart

:::image type="content" source="../../images/cone-column-chart.png" alt-text="A 3D cone column chart showing quarterly performance with cone-shaped columns.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Quarter", "Target", "Achieved"],
    ["Q1", 100000, 95000],
    ["Q2", 120000, 125000],
    ["Q3", 130000, 128000],
    ["Q4", 150000, 158000]
  ];
  const dataRange = sheet.getRange("A1:C5");
  dataRange.setValues(data);
  
  // Create cone column chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.coneCol,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Quarterly Performance (3D Cone)");
  
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

### Cylinder column chart

:::image type="content" source="../../images/cylinder-column-chart.png" alt-text="A 3D cylinder column chart displaying product sales with cylindrical columns.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Product", "Units Sold"],
    ["Laptops", 850],
    ["Tablets", 720],
    ["Phones", 640],
    ["Monitors", 580]
  ];
  const dataRange = sheet.getRange("A1:B5");
  dataRange.setValues(data);
  
  // Create cylinder column chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.cylinderCol,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Product Sales (3D Cylinder)");
  
  chart.getLegend().setVisible(false);
}
```

### Pyramid column chart

:::image type="content" source="../../images/pyramid-colum-chart.png" alt-text="A 3D pyramid column chart showing organizational hierarchy with pyramid-shaped columns.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Level", "Headcount"],
    ["Executive", 5],
    ["Senior Management", 25],
    ["Middle Management", 120],
    ["Staff", 450]
  ];
  const dataRange = sheet.getRange("A1:B5");
  dataRange.setValues(data);
  
  // Create pyramid column chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.pyramidCol,
    dataRange
  );
  chart.setPosition("D1");
  chart.getTitle().setText("Organizational Hierarchy (3D Pyramid)");
  
  chart.getLegend().setVisible(false);
}
```

## Working with chart elements

This sample demonstrates how to customize various chart elements that apply across all chart types.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  
  // Add sample data.
  const data = [
    ["Month", "Revenue", "Expenses"],
    ["Jan", 45000, 32000],
    ["Feb", 52000, 35000],
    ["Mar", 48000, 33000],
    ["Apr", 61000, 38000]
  ];
  const dataRange = sheet.getRange("A1:C5");
  dataRange.setValues(data);
  
  // Create chart.
  const chart = sheet.addChart(
    ExcelScript.ChartType.columnClustered,
    dataRange
  );
  chart.setPosition("E1");
  
  // Customize chart title.
  chart.getTitle().setText("Monthly Financial Overview");
  const chartTitle = chart.getTitle();
  chartTitle.getFormat().getFont().setSize(16);
  chartTitle.getFormat().getFont().setBold(true);
  chartTitle.getFormat().getFont().setColor("#2C3E50");
  
  // Customize legend.
  const legend = chart.getLegend();
  legend.setPosition(ExcelScript.ChartLegendPosition.bottom);
  legend.getFormat().getFont().setSize(10);
  legend.setVisible(true);
  
  // Customize axes.
  const valueAxis = chart.getAxes().getValueAxis();
  valueAxis.setDisplayUnit(ExcelScript.ChartAxisDisplayUnit.thousands);
  valueAxis.getMajorGridlines().getFormat().getLine().setColor("#D3D3D3");
  valueAxis.getTitle().setText("Amount (in thousands)");
  
  const categoryAxis = chart.getAxes().getCategoryAxis();
  categoryAxis.getTitle().setText("Month");
  
  // Customize series.
  const series = chart.getSeries();
  series[0].getFormat().getFill().setSolidColor("#3498DB"); // Revenue - Blue.
  series[1].getFormat().getFill().setSolidColor("#E74C3C"); // Expenses - Red.
  
  // Set chart size.
  chart.setHeight(300);
  chart.setWidth(500);
}
```

## See also

- [Chart object (Office Scripts API)](https://learn.microsoft.com/javascript/api/office-scripts/excelscript/excelscript.chart)
- [ChartType enum (Office Scripts API)](https://learn.microsoft.com/javascript/api/office-scripts/excelscript/excelscript.charttype)
- [Office Scripts samples and scenarios](samples-overview.md)
- [Tutorial: Create and format an Excel table](../../tutorials/excel-tutorial.md)
