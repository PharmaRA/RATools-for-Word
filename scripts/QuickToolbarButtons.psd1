@{
    # =============================================
    # 快捷悬浮工具栏按钮定义单一事实源
    # 消费方：scripts/Sync-QuickToolbarForm.ps1（生成 frm/frx）
    #         tests/QuickToolbar.Tests.ps1（推导全部期望值）
    # 修改按钮/布局后需重跑 Sync 脚本再构建。
    # Kind: custom=项目 PNG 图标 / mso=Word 内置图标(idMso) / mso-image=GetImageMso 位图
    # Dispatch: 点击后的行为契约（测试断言分发结果）
    # =============================================
    Layout = @{
        ToolbarWidth  = 134
        ToolbarHeight = 168
        ButtonWidth   = 26
        ButtonHeight  = 24
        RowStep       = 25
        ColumnGap     = 3
        ButtonTop     = 3
        GroupGap      = 4
        PictureSize   = 20
        # 从这些行号起各累加一次组间距
        GroupGapRows  = @(2, 4)
    }
    Buttons = @(
        @{ Name = "btnHeading1"; Caption = "编号标题 1"; Tip = "应用标题1-F样式"; Icon = "H1"; Kind = "custom"; Action = "style_heading1"; Row = 0; Column = 0; ColumnCount = 4; Dispatch = "style:标题1-F" }
        @{ Name = "btnHeading2"; Caption = "编号标题 2"; Tip = "应用标题2-F样式"; Icon = "H2"; Kind = "custom"; Action = "style_heading2"; Row = 0; Column = 1; ColumnCount = 4; Dispatch = "style:标题2-F" }
        @{ Name = "btnHeading3"; Caption = "编号标题 3"; Tip = "应用标题3-F样式"; Icon = "H3"; Kind = "custom"; Action = "style_heading3"; Row = 0; Column = 2; ColumnCount = 4; Dispatch = "style:标题3-F" }
        @{ Name = "btnHeading4"; Caption = "编号标题 4"; Tip = "应用标题4-F样式"; Icon = "H4"; Kind = "custom"; Action = "style_heading4"; Row = 0; Column = 3; ColumnCount = 4; Dispatch = "style:标题4-F" }
        @{ Name = "btnUnnumberedHeading1"; Caption = "无编号标题 1"; Tip = "应用无编号标题1-F样式"; Icon = "UNH1"; Kind = "custom"; Action = "style_unnumbered_heading1"; Row = 1; Column = 0; ColumnCount = 4; Dispatch = "style:无编号标题1-F" }
        @{ Name = "btnUnnumberedHeading2"; Caption = "无编号标题 2"; Tip = "应用无编号标题2-F样式"; Icon = "UNH2"; Kind = "custom"; Action = "style_unnumbered_heading2"; Row = 1; Column = 1; ColumnCount = 4; Dispatch = "style:无编号标题2-F" }
        @{ Name = "btnUnnumberedHeading3"; Caption = "无编号标题 3"; Tip = "应用无编号标题3-F样式"; Icon = "UNH3"; Kind = "custom"; Action = "style_unnumbered_heading3"; Row = 1; Column = 2; ColumnCount = 4; Dispatch = "style:无编号标题3-F" }
        @{ Name = "btnUnnumberedHeading4"; Caption = "无编号标题 4"; Tip = "应用无编号标题4-F样式"; Icon = "UNH4"; Kind = "custom"; Action = "style_unnumbered_heading4"; Row = 1; Column = 3; ColumnCount = 4; Dispatch = "style:无编号标题4-F" }
        @{ Name = "btnFormatPainter"; Caption = "格式刷"; Tip = "复制并应用所选内容的格式"; Icon = "FormatPainter"; Kind = "mso"; Action = "format_painter"; Row = 2; Column = 0; ColumnCount = 4; Dispatch = "mso:FormatPainter" }
        @{ Name = "btnParagraphSettings"; Caption = "段落设置"; Tip = "打开 Word 段落设置对话框"; Icon = "ParagraphDialog"; Kind = "mso"; Action = "paragraph_settings"; Row = 2; Column = 1; ColumnCount = 4; Dispatch = "mso:ParagraphDialog" }
        @{ Name = "btnPageBreakBefore"; Caption = "段前分页"; Tip = "切换选中段落的段前分页属性"; Icon = "PageBreakBefore"; Kind = "custom"; Action = "page_break_before"; Row = 2; Column = 2; ColumnCount = 4; Dispatch = "page_break_before" }
        @{ Name = "btnTableTitle"; Caption = "表标题"; Tip = "应用表标题-F样式"; Icon = "TableTitle"; Kind = "custom"; Action = "style_table_title"; Row = 2; Column = 3; ColumnCount = 4; Dispatch = "style:表标题-F" }
        @{ Name = "btnFigureTitle"; Caption = "图标题"; Tip = "应用图标题-F样式"; Icon = "FigureTitle"; Kind = "custom"; Action = "style_figure_title"; Row = 3; Column = 0; ColumnCount = 4; Dispatch = "style:图标题-F" }
        @{ Name = "btnAutoFitTable"; Caption = "根据窗口自动调整表格"; Tip = "将光标所在表格根据窗口宽度自动调整"; Icon = "TableAutoFitWindow"; Kind = "mso-image"; Action = "autofit_table"; Row = 3; Column = 1; ColumnCount = 4; Dispatch = "autofit_table" }
        @{ Name = "btnInsertCrossReference"; Caption = "插入交叉引用"; Tip = "打开插入交叉引用对话框"; Icon = "CrossReferenceInsert"; Kind = "mso"; Action = "insert_cross_reference"; Row = 3; Column = 2; ColumnCount = 4; Dispatch = "mso:CrossReferenceInsert" }
        @{ Name = "btnUpdateFields"; Caption = "更新域"; Tip = "更新当前选区或光标所在位置的域"; Icon = "FieldsUpdate"; Kind = "mso"; Action = "update_fields"; Row = 3; Column = 3; ColumnCount = 4; Dispatch = "mso:FieldsUpdate" }
        @{ Name = "btnHyperlinksFieldsBlue"; Caption = "超链接和域批量设置为蓝色"; Tip = "运行超链接和域批量设置为蓝色宏"; Icon = "HyperlinkInsert"; Kind = "mso"; Action = "hyperlinks_fields_blue"; Row = 4; Column = 0; ColumnCount = 3; Dispatch = "hyperlinks_fields_blue" }
        @{ Name = "btnAcceptRevisionsComments"; Caption = "接受修订并删除批注"; Tip = "运行接受修订并删除批注宏"; Icon = "ReviewAcceptChange"; Kind = "mso"; Action = "accept_revisions_comments"; Row = 4; Column = 1; ColumnCount = 3; Dispatch = "accept_revisions_comments" }
        @{ Name = "btnDetectHighlights"; Caption = "检测高亮内容"; Tip = "运行检测高亮内容宏"; Icon = "TextHighlightColorPicker"; Kind = "mso"; Action = "detect_highlights"; Row = 4; Column = 2; ColumnCount = 3; Dispatch = "detect_highlights" }
    )
}
