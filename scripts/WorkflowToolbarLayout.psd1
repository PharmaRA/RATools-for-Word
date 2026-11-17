@{
    Writing = @{
        Caption = 'RATools 撰写工具'; Columns = 4; Width = 392
        Rows = @(
            @('章节结构', '标题 1|style:标题1-F', '标题 2|style:标题2-F', '标题 3|style:标题3-F', '标题 4|style:标题4-F'),
            @('', '无编号 1|style:无编号标题1-F', '无编号 2|style:无编号标题2-F', '无编号 3|style:无编号标题3-F', '无编号 4|style:无编号标题4-F'),
            @('正文样式', '正文|style:正文-F', '正文无缩进|style:正文无缩进-F', '正文无间距|style:正文无间距-F', '更多样式|menu:styles'),
            @('列表与文字', '编号列表|style:编号列表-F', '项目符号|style:项目符号列表-F', '插入脚注|mso:FootnoteInsert', '更多文字|menu:text'),
            @('段落格式', '格式刷|mso:FormatPainter', '清除格式|mso:ClearFormatting', '段落设置…|mso:ParagraphDialog', '段前分页|page_break_before'),
            @('表格与页面', '当前表适宽|autofit_table', '选中文字标蓝|text_blue', '页面设置…|mso:PageSetupDialog', '更多排版|menu:layout')
        )
        Menus = @{
            styles = @(
                '— 编号标题 —|', '标题 5|style:标题5-F', '标题 6|style:标题6-F', '标题 7|style:标题7-F', '标题 8|style:标题8-F', '标题 9|style:标题9-F',
                '— 无编号标题 —|', '无编号 5|style:无编号标题5-F', '无编号 6|style:无编号标题6-F', '无编号 7|style:无编号标题7-F', '无编号 8|style:无编号标题8-F', '无编号 9|style:无编号标题9-F',
                '— 其他样式 —|', '标题居中|style:标题居中-F', '标题左对齐|style:标题左对齐-F', '脚注样式|style:脚注-F', '参考文献列表|style:参考文献列表-F', '样式窗格|style_pane'
            )
            text = @('上标|mso:Superscript', '下标|mso:Subscript', '标准化科学术语下标|normalize_terms', '提取缩略语|extract_abbreviations')
            layout = @(
                '— 局部操作 —|', '插入分页符|mso:PageBreakInsertWord', '插入分节符|mso:SectionBreakInsert', '页眉样式|style:页眉-F', '页脚样式|style:页脚-F',
                '— 整篇处理 —|', '当前文档表格统一适宽|autofit_all_tables', '页眉页脚链接到前一节|link_sections', '批量设置页边距〔各节四边 2.54 cm〕|batch_margins'
            )
        }
    }
}
