# goppt

A Go library for working with PowerPoint presentations. It handles reading, writing, and rendering `.pptx` files with no external dependencies.

## Overview

goppt implements the OOXML specification used by PowerPoint 2007 and later. You can build presentations from scratch, parse existing files, and export slides to PNG images — all from pure Go code.

The slide renderer is designed to produce output that closely matches what PowerPoint itself generates. It handles the tricky parts: text layout, font metrics, CJK line-breaking rules, shape geometry, and chart drawing.

## Getting Started

```bash
go get github.com/casibase/goppt
```

Requires Go 1.18 or later.

## Building a Presentation

```go
package main

import (
    "log"
    ppt "github.com/casibase/goppt"
)

func main() {
    p := ppt.New()
    p.GetDocumentProperties().Title = "Quarterly Review"
    p.GetDocumentProperties().Creator = "goppt"

    slide := p.GetActiveSlide()

    heading := slide.CreateRichTextShape()
    heading.SetOffsetX(500000).SetOffsetY(300000)
    heading.SetWidth(8000000).SetHeight(1000000)
    run := heading.CreateTextRun("Q3 Results")
    run.GetFont().SetSize(32).SetBold(true).SetColor(ppt.ColorBlue)

    body := slide.CreateRichTextShape()
    body.SetOffsetX(500000).SetOffsetY(1500000)
    body.SetWidth(8000000).SetHeight(600000)
    body.CreateTextRun("Revenue up 18% year over year")

    next := p.CreateSlide()
    chart := next.CreateChartShape()
    chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
    chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
    chart.GetTitle().SetText("Revenue by Quarter")

    bars := ppt.NewBarChart()
    bars.AddSeries(ppt.NewChartSeriesOrdered(
        "Revenue",
        []string{"Q1", "Q2", "Q3", "Q4"},
        []float64{120, 180, 150, 210},
    ))
    chart.GetPlotArea().SetType(bars)

    w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
    if err := w.(*ppt.PPTXWriter).Save("review.pptx"); err != nil {
        log.Fatal(err)
    }
}
```

## Reading an Existing File

```go
reader := &ppt.PPTXReader{}
pres, err := reader.Read("existing.pptx")
if err != nil {
    log.Fatal(err)
}

for i, s := range pres.GetAllSlides() {
    fmt.Printf("slide %d — %d shapes\n", i+1, len(s.GetShapes()))
}
```

## Writing to a Buffer

```go
var buf bytes.Buffer
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.WriteTo(&buf)
```

## Capabilities

**Text and typography**
- Rich text runs with font family, size, color, bold, italic, underline, and strikethrough
- Text boxes with auto-fit, auto-shrink, and overflow modes
- Bullet lists (character and numeric)

**Shapes**
- Auto shapes: rectangles, ellipses, triangles, arrows, stars, callouts, and more
- Line shapes with configurable style and color
- Group shapes and placeholder shapes

**Images**
- Embed PNG, JPEG, GIF, BMP, or SVG from a file path or raw bytes
- Rotation, flip, and group transform support

**Charts**
- Bar, Bar3D, Line, Area, Pie, Pie3D, Doughnut, Scatter, Radar

**Tables**
- Cell-level formatting, fills, and borders

**Slide options**
- Solid and gradient backgrounds
- Speaker notes
- Comments with author metadata
- Basic animation grouping
- Document and custom properties
- Slide sizes: 4:3, 16:9, 16:10, A4, Letter, or custom

## Renderer

The built-in renderer exports slides to PNG. Key behaviors:

- Uses two hinting modes (HintingNone for layout metrics, HintingFull for glyph rendering) to match PowerPoint's DirectWrite measurements
- Implements kinsoku line-breaking rules for CJK text
- Renders shape fills, strokes, drop shadows, and custom geometry paths
- Draws all supported chart types with axes and series colors

## Reference

See [API.md](API.md) for the complete API reference.

## License

MIT
