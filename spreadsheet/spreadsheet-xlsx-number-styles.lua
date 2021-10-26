local M = {}
local built_in_styles = {
["0"] = 'General',
["1"] = '0',
["2"] = '0.00',
["3"] = '#,##0',
["4"] = '#,##0.00',
["5"] = '$#,##0,\\-$#,##0';
["6"] = '$#,##0,[Red]\\-$#,##0';
["7"] = '$#,##0.00,\\-$#,##0.00';
["8"] = '$#,##0.00,[Red]\\-$#,##0.00';
["9"] = '0%',
["10"] = '0.00%',
["11"] = '0.00E+00',
["12"] = '# ?/?',
["13"] = '# ??/??',
["14"] = 'mm-dd-yy',
["15"] = 'd-mmm-yy',
["16"] = 'd-mmm',
["17"] = 'mmm-yy',
["18"] = 'h:mm AM/PM',
["19"] = 'h:mm:ss AM/PM',
["20"] = 'h:mm',
["21"] = 'h:mm:ss',
["22"] = 'm/d/yy h:mm',
["37"] = '#,##0 ,(#,##0)';
["38"] = '#,##0 ,[Red](#,##0)';
["39"] = '#,##0.00,(#,##0.00)';
["40"] = '#,##0.00,[Red](#,##0.00)';
["44"] = '_("$"* #,##0.00_),_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
["45"] = 'mm:ss',
["46"] = '[h]:mm:ss',
["47"] = 'mmss.0',
["48"] = '##0.0E+0',
["49"] = '@',
["27"] = '[$-404]e/m/d',
["30"] = 'm/d/yy',
["36"] = '[$-404]e/m/d',
["50"] = '[$-404]e/m/d',
["57"] = '[$-404]e/m/d',
["59"] = 't0',
["60"] = 't0.00',
["61"] = 't#,##0',
["62"] = 't#,##0.00',
["67"] = 't0%',
["68"] = 't0.00%',
["69"] = 't# ?/?',
["70"] = 't# ??/??'
}

function load_styles(dom)
  local number_formats = built_in_styles
  local cell_styles = {}
  -- load number formats
  for _, numfmt in ipairs(dom:query_selector("numFmt")) do
    numFmtId="165" formatCode="dd/MM/yyyy"
    local id, code = numfmt:get_attribute("numfmtid"), numfmt:get_attribute("formatcode")
    number_formats[id] = code
  end
  for i, xf in ipairs(dom:query_selector("cellXfs xf")) do 
    -- we should add more formats in the future
    local num_format = xf:get_attribute("numFmtId")
  end
  local styles = {dom = dom, number_formats = number_formats}
  return styles
end

M.load_styles = load_styles
return M
