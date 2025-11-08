# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-11-08

### Added
- Initial release of html2excel-richtext
- Complete rich text support (bold, italic, underline, colors, fonts, sizes)
- Pure Java CSS parser with high fault tolerance
- Cell background color mapping
- Hyperlink auto-extraction and conversion
- Image download and embedding support
- Font and style caching for performance
- Configurable converter settings via ConverterConfig
- Comprehensive unit tests
- Clean architecture following SOLID principles

### Features
- Support for HTML tags: `<b>`, `<strong>`, `<i>`, `<em>`, `<u>`, `<a>`, `<p>`, `<div>`, etc.
- Support for CSS properties: `color`, `background-color`, `font-family`, `font-size`, `font-weight`, `font-style`, `text-decoration`
- Support for multiple color formats: #hex, rgb(), named colors
- Auto-truncation for long text (>32,767 characters)
- Thread-safe caching with ConcurrentHashMap

### Architecture
- **HtmlToExcelConverter**: Main facade class
- **ConverterConfig**: Configuration with builder pattern
- **Parser**: CssParser, ColorParser, HtmlTraverser
- **Cache**: FontCache, StyleCache
- **Builder**: FontBuilder
- **Handler**: BackgroundHandler, HyperlinkHandler, ImageHandler

### Documentation
- Comprehensive README with examples
- JavaDoc comments for all public APIs
- Apache 2.0 License

[1.0.0]: https://github.com/fivefish130/html2excel-richtext/releases/tag/v1.0.0
