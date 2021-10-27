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
    local id, code = numfmt:get_attribute("numfmtid"), numfmt:get_attribute("formatcode")
    number_formats[id] = code
  end
  local cell_styles = {}
  for i, xf in ipairs(dom:query_selector("cellXfs xf")) do 
    -- we should add more formats in the future
    local num_format = xf:get_attribute("numfmtid")
    cell_styles[#cell_styles+1] = {
      num_format = num_format
    }
  end
  local styles = {dom = dom, number_formats = number_formats, cell_styles = cell_styles}
  return styles
end

local function get_style(styles, orig_id)
  local id = tonumber(orig_id)
  if id then 
    --- get table with cell styles
    local cell_styles = styles.cell_styles or {}
    -- get the right id number
    id = id + 1
    return cell_styles[id]
  end
  return nil, "cannot load style"
end

local function get_number_format(styles, style)
  local number_formats = styles.number_formats or {}
  local style = style or {}
  -- default style is "1"
  local id = style.num_format or "1"
  return number_formats[id]
end

local function convert_date(num)
  -- Excel dates start at year 1900. This function converts 
  -- it to timestamp suitable for os.date function
  return math.floor((num - 25569) * 86400)
end

-- convert Excel format to string suitable for string.format
local function get_string_format(num_format)
  -- general formatting string used by default
  if num_format == "General" then return "%s", false end
  -- handle numbers
  if num_format:match("[%#0]") then
    -- number formats
    -- at the moment we just print float or integer
    if num_format:match("[%.%,]") then
      return num_format:gsub("[0%#%.%,]+", "%%f")
    else
      return num_format:gsub("[0%#%.%]+", "%%i")
    end
  else
    -- TODO: we just assume that other than general and number, format is date
    -- which is obviously wrong. it should be fixed in the future
    -- fix AM/PM
    num_format = num_format:gsub("AM%/PM", "%%p")
    -- fix months
    num_format = num_format:gsub("([^m])m([^m])", "%1%%m%2"):gsub("MM", "%%m"):gsub(
    "mmmmm", "%%b"):gsub( "mmmm", "%%B"):gsub( "mmm", "%%b"):gsub("mm", "%%m")
    -- fix minutes
    num_format = num_format:gsub("h(.)%%m", "h%1%%M"):gsub("%%m(.)s", "%%M%1s")
    -- fix hours
    if num_format:match("%%p") then -- use am/pm
      num_format = num_format:gsub("[h]+", "%%I")
    else
      num_format = num_format:gsub("[h]+", "%%H")
    end
    -- fix seconds
    num_format = num_format:gsub("ss", "%%S"):gsub("s", "%%S")
    -- fix days
    num_format = num_format:gsub("([^d])d([^d])", "%1%%d%2"):gsub("dddd", "%%A"):gsub(
    "ddd", "%%a"):gsub("dd", "%%d")
    -- fix years
    num_format = num_format:gsub("yyyy", "%%Y"):gsub("yy", "%%y")
    return num_format, true
  end
  -- return string by default
  return "%s"
end

local function apply_number_format(num_format, value)
  local num = tonumber(value)
  local format, is_date = get_string_format(num_format)
  if num then
    if is_date then
      local timestamp = convert_date(num)
      -- print(value, num, timestamp)
      -- print(os.date("%c", timestamp))
      return os.date(format, timestamp)
    else
      return string.format(format, num)
    end
  end
  return value
end

M.load_styles = load_styles
M.apply_number_format = apply_number_format
M.get_style   = get_style
M.get_number_format = get_number_format
return M
