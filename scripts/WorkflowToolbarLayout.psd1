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
    References = @{
        Caption = 'RATools 图表引用工具'; Columns = 2; Width = 432
        Rows = @(
            @('建立题注与引用', '生成题注与引用…|caption_pair', '插入题注…|mso:CaptionInsert'),
            @('', '插入交叉引用…|mso:CrossReferenceInsert', '恢复题注标签|restore_caption_labels'),
            @('图表格式', '表标题样式|style:表标题-F', '图标题样式|style:图标题-F'),
            @('', '当前表适宽|autofit_table', '更多图表工具|menu:table_tools'),
            @('复制与维护', '粘贴并重连引用|paste_relink', '更新当前范围域|mso:FieldsUpdate'),
            @('', '引用与链接标蓝|links_fields_blue', '保护引用域格式|protect_field_format'),
            @('图表目录', '插入图表目录…|mso:TableOfFiguresInsert', '更新图表目录…|mso:TableOfFiguresUpdate')
        )
        Menus = @{
            table_tools = @(
                '— 表格文字 —|', '表头左对齐|style:表头左对齐-F', '表头居中|style:表头居中-F', '表头右对齐|style:表头右对齐-F',
                '表格正文左对齐|style:表格文本左对齐-F', '表格正文居中|style:表格文本居中-F', '表格正文右对齐|style:表格文本右对齐-F',
                '无间距|style:表格文本无间距-F', '表格注释|style:表格注释-F',
                '— 表格结构 —|', '插入表格|mso:TableInsertGallery', '边框|mso:TableBordersMenu', '重复标题行|mso:TableRepeatHeaderRows', '拆分表格|mso:TableSplitTable',
                '— 图片 —|', '图片段落样式|style:图片-F', '嵌入型环绕|mso:TextWrappingInLineWithText'
            )
        }
        Tips = @{
            'caption_pair' = '光标须位于正文中，并在表格与现有域之外；先取消文字选区，不自动插入图片或表格。'
            'style:表标题-F' = '应用表标题-F样式；只设置样式，自动编号通过题注功能建立。'
            'style:图标题-F' = '应用图标题-F样式；只设置样式，自动编号通过题注功能建立。'
            'paste_relink' = '复制后在目标位置点击，无需先粘贴；此动作包含一次粘贴和引用重连。'
            'restore_caption_labels' = '恢复 Word 题注类型入口，不表示恢复已删除的引用目标或书签。'
            'links_fields_blue' = '沿用当前文档批量处理及排除规则；不是选区上色，也不是引用检查报告。'
            'mso:FieldsUpdate' = '更新当前选区或所在域，不表示全篇所有文字区域更新。'
            'protect_field_format' = '保留 REF、PAGEREF 域格式；有选区时处理选区，无选区时处理当前正文域集合。'
        }
    }
    Final = @{
        Caption = 'RATools 定稿工具'; Columns = 2; Width = 432
        Rows = @(
            @('交付前检查', '检测 clean 版本…|detect_clean_versions', '检测高亮内容…|detect_highlights'),
            @('目录与引用', '更新章节目录…|mso:TableOfContentsUpdate', '更新图表目录…|mso:TableOfFiguresUpdate'),
            @('', '更新当前范围域|mso:FieldsUpdate', '保护引用域格式|protect_field_format'),
            @('清稿与编号固定', '接受修订并删除批注…|accept_revisions_comments', '标题编号转文本|heading_numbers_text'),
            @('保存与交付', '保存当前文档|save_current', '导出 PDF…|export_pdf'),
            @('', '重命名当前文件…|rename_current', '打开所在文件夹|open_folder')
        )
        Menus = @{}
        Tips = @{
            'detect_clean_versions' = '沿用当前文档、多文件和文件夹模式；高亮不纳入 clean 判定。'
            'mso:TableOfContentsUpdate' = '沿用 Word 当前章节目录的定位和更新方式。'
            'mso:TableOfFiguresUpdate' = '沿用 Word 当前图表目录的定位和更新方式。'
            'mso:FieldsUpdate' = '更新当前选区或所在域；不表示全篇所有文字区域更新。'
            'protect_field_format' = '保留 REF、PAGEREF 域格式；有选区时处理选区，无选区时处理当前正文域集合。'
            'heading_numbers_text' = '处理当前文档正文中的符合条件标题；适合最终副本，不自动另存或保存。'
            'save_current' = '保存当前活动文档；取消或失败后保留实际未保存状态。'
            'export_pdf' = '保留现有处理模式及 PDF 选项；未保存编辑可能进入 PDF，Word 原件不因此自动保存。'
            'accept_revisions_comments' = '沿用现有处理模式、关闭修订跟踪及保存规则。'
        }
    }
}
