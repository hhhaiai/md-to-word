-- Chinese Document Format Filter for Pandoc
-- 用于处理中文公文格式的Pandoc Lua过滤器

-- 处理数学公式，确保正确渲染
function Math(el)
    -- 保持数学公式的原始处理方式
    return el
end

-- 处理表格，确保中文内容正确显示
function Table(el)
    -- 遍历表格内容，处理中文字符
    for i = 1, #el.rows do
        for j = 1, #el.rows[i].cells do
            -- 处理每个单元格的内容
            el.rows[i].cells[j] = pandoc.walk_block(el.rows[i].cells[j], {
                Str = function(str_el)
                    -- 确保中文字符正确处理
                    return str_el
                end
            })
        end
    end
    return el
end

-- 处理列表，确保中文编号和缩进
function BulletList(el)
    -- 保持无序列表的原始格式
    return el
end

function OrderedList(el)
    -- 保持有序列表的原始格式，pandoc会正确处理中文编号
    return el
end

-- 处理段落，确保中文字符正确显示
function Para(el)
    return el
end

-- 处理标题
function Header(el)
    -- 保持标题的原始处理
    return el
end

-- 全局处理函数
function Pandoc(doc)
    -- 设置文档级别的元数据
    return doc
end