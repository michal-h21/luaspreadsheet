--- XSLT reader module
local M = {}
local zip = require "zip"
local domobject = require "luaxml-domobject"
local log = require "spreadsheet.spreadsheet-log"
local ranges = require "spreadsheet.spreadsheet-ranges"


--- @type Xlsx object
local Xlsx = {}
Xlsx.__index = Xlsx

--- @type Sheet object
local Sheet ={}
Sheet.__index = Sheet

--- load the xlsx file
-- @return Xlsx object
local function load(filename)
  -- to add support for --reader option of TeX, we may use the following trick
  local f = io.open(filename, "r")
  if not f then return nil, "File ".. filename .. " not found" end
  f:close()
  local xlsx_obj = setmetatable({}, Xlsx)
  local status, msg = xlsx_obj:load(filename)
  if not status then
    return nil, msg
  end
  return xlsx_obj
end

-- helper function which simplifies DOM loading from the zip file
local function load_zip_xml(zipfile, path)
  local f = zipfile:open(path)
  if not f then
    return nil, "Cannot find file in the XLSX archive: ".. path
  end
  local text = f:read("*all")
  local dom = domobject.parse(text)
  if not dom then
    return nil, "Error in parsing XML file: " .. path
  end
  return dom
end


function Xlsx:load(filename)
  log.info("Loading "..filename)
  local zip_file = zip.open(filename)
  self.file = zip_file
  self.saved_sheets = {}
  -- load file with pointer to other files
  local content_types, msg = self:load_zip_xml("[Content_Types].xml")
  if not content_types then
    return nil, msg
  end
  -- load default relationships
  self:load_relationships("/_rels/.rels", "")
  self.content_types = content_types
  -- load workbook, styles and shared_strings, save worksheets
  for _, override in ipairs(content_types:query_selector("Override")) do
    local content_type = override:get_attribute("contenttype")
    local part_name = override:get_attribute("partname")
    if content_type:match("sheet.main") then
      self:load_workbook(part_name)
    elseif content_type:match("styles") then
      self:load_styles(part_name)
    elseif content_type:match("sharedStrings") then
      self:load_shared_strings(part_name)
    -- elseif content_type:match("worksheet") then
      -- for k,v in pairs(override._attr) do print("worksheet", k,v) end
    elseif content_type:match("relationships") then
      self:load_relationships(part_name)
    end
    -- print(content_type, part_name)
    -- print(override:get_attribute("PartName"), override:get_attribute("ContentType"))
  end
  return true
end

--- Parse XML file from the xlsx archive to DOM
function Xlsx:load_zip_xml(filename)
  -- we must remove the slash at the beginning, the zip library
  -- would not find the file otherwise
  local filename = filename:gsub("^/", "")
  log.info("loading: ".. filename)
  local zip_file = self.file
  local dom, msg = load_zip_xml(zip_file,filename)
  -- if not dom then 
    -- log.error("Cannot parse XML from " .. filename)
  -- end
  return dom, msg
end

function Xlsx:get_directory(filename)
  return filename:match("(.-)[^/]+$")
end

function Xlsx:load_file_relationships(filename)
  local directory = self:get_directory(filename)
  local basename = filename:match("([^%/]+)$")
  self:load_relationships(directory .. "_rels/" .. basename .. ".rels" )
  self:load_relationships(directory .. "_rels/" .. basename .. ".rels" , "/" .. filename)
end

function Xlsx:load_workbook(filename)
  local workbook,msg = self:load_zip_xml(filename)
  -- because the paths in _rels table are relative to the current 
  -- directory, we must save the current path, which will be then 
  -- used in the file loading
  local directory = self:get_directory(filename)
  workbook.directory = directory
  log.info("workbook path:".. directory)
  self:load_file_relationships(filename)
  local sheets = {}
  -- save the worksheet names and identifiers
  for _, sheet in ipairs(workbook:query_selector("sheet")) do
    sheets[#sheets+1] = {
      name = sheet:get_attribute("name"),
      sheetid = sheet:get_attribute("sheetid"),
      id = sheet:get_attribute("r:id")
    }
  end
  self.sheets = sheets
  self.workbook = workbook
end

--- normalize relative paths
function Xlsx:normalize_path(path)
  local parts = {}
  -- this regexp doesn't keep the leading slash, but it doesn't matter
  -- because we need to get rid of it anyway
  for part in path:gmatch("([^%/]+)") do
    -- remove the up directory when we find the ".." part
    if part == ".." then
      parts[#parts] = nil
    else
      parts[#parts+1] = part
    end
  end
  return table.concat(parts, "/")
end

function Xlsx:load_relationships(filename, path)
  -- each relationship file correspond to the parent directory
  -- so we must find that corresponding dir
  local path = path or filename:match("(.-)_rels")
  local relationships = self.relationships or {}
  local current = {}
  local dom = self:load_zip_xml(filename)
  if not dom then return false, "Cannot load relationships file" .. filename end
  for _, el in ipairs(dom:query_selector("Relationship")) do
    local id, target, schema = el:get_attribute("id"), el:get_attribute("target"), el:get_attribute("type")
    -- we must construct full path to the target file. but exclude hyperlinks
    local fulltarget = target
    if not schema:match("hyperlink$") then
      fulltarget = path .. target
    end
    fulltarget = self:normalize_path(fulltarget)
    current[id] = {target = fulltarget, type = schema}
  end
  relationships[path] = current
  self.relationships = relationships
end

function Xlsx:load_styles(filename)
  local  dom = self:load_zip_xml(filename)
  -- for _, cell in ipairs(dom:query_selector("cellStyle")) do
    -- print(cell:serialize())
  -- end
  self.styles = dom
end

function Xlsx:load_shared_strings(filename)
  local string_dom = self:load_zip_xml(filename)
  local shared_strings = {}
  local pos = 0
  for _, si in ipairs(string_dom:query_selector("si")) do
    shared_strings[pos] = si
    pos = pos+1
  end
  -- we must correct the pos variable for logging
  if pos > 0 then pos = pos - 1 end
  log.info("Loaded ".. pos .." shared strings")
  -- print(string_dom:serialize())
  self.shared_strings = shared_strings
end

function Xlsx:find_file_by_id(rid,directory)
  local workbook = self.workbook
  -- the path to sheet file is saved in the relationships table
  -- there are several of such tables, for different directories
  -- we must retrieve the one for the directory where the workbook
  -- file lies
  local relationships = self.relationships or {}
  local directory = directory or  workbook.directory
  -- the relationship directories were saved with leading slashes, we must
  -- add it if it was removed during path normalization
  directory = directory:match("^/") and directory or "/"..directory
  local rel_table = relationships[directory] or relationships[""] or {}
  local relation_dest = rel_table[rid]
  if relation_dest then
    local target = relation_dest.target
    log.info("Found file for id " .. rid.. ": ".. target)
    return target
  end
  local msg = "Cannot find file id " .. rid
  log.error(msg)
  return nil, msg
end

function Xlsx:load_sheet(name)
  local sheet_obj = self.saved_sheets[name]
  if not sheet_obj then
    local dom = self:load_zip_xml(name)
    sheet_obj = setmetatable({}, Sheet)
    sheet_obj:set_parent(self,name)
    sheet_obj:load_dom(name, dom)
    self.saved_sheets[name] = sheet_obj
  end
  return sheet_obj
end

function Xlsx:find_sheet_id(name)
  local name = name or 1
  local attr = "name"
  if type(name) == "number" then
    attr = "sheetid"
    -- the attribute value is string, so we must convert the name
    -- to string to assure they match
  end
  name = tostring(name)
  for _, sheet in ipairs(self.sheets) do
    if sheet[attr] == name then
      return sheet.id
    end
  end
end

--- Retrieve sheet from workbook.
-- It can be referenced either by name, or id number.
function Xlsx:get_sheet(
  name -- string or number
)
  local rid = self:find_sheet_id(name)
  if rid then
    local target, msg = self:find_file_by_id(rid)
    if target then
      return self:load_sheet(target)
    end
    return nil, msg
  end
  local msg = "Cannot find sheet " .. name
  log.error(msg)
  return nil, msg
end

function Sheet:set_parent(parent,filename)
  self.parent = parent
  self.relationships = parent.relationships
  self.directory = parent:get_directory(filename)
  self.load_zip_xml = parent.load_zip_xml
  self.find_file_by_id = parent.find_file_by_id
  self.shared_strings = parent.shared_strings
  self.file = parent.file
  self.table = {}
end

function Sheet:get_parent()
  return self.parent
end

function Sheet:load_dom(name, dom)
  local rows = #dom:query_selector("row")
  local parent = self:get_parent()
  parent:load_file_relationships(name)
  self:save_dimensions(dom)
  self:load_merge_cells(dom)
  self:load_named_ranges(dom)
  self:load_links(dom, name)
  self:load_columns(dom)
  -- make table with all data
  self:process_rows(dom)
  -- todo: add merged cells and links to the table
  self:save_links()
  -- xxx("sheetData")
  -- xxx("dimension")
  return dom
end

function Sheet:load_merge_cells(dom)
  local merge_cell = dom:query_selector("mergeCell")
  -- merge cells are included in the <mergeCell> element
  local merge_cells = self.merge_cells or {}
  for _, merge in ipairs(merge_cell) do
    -- ref are in the A1:B2 form
    -- we will save them in a hash table, indexed by the first reference, because
    -- it is possible to retrieve them quickly, as all cells have reference attribute
    local ref = merge:get_attribute("ref")
    local first_ref = ref:match("([^%:]+)")
    merge_cells[first_ref] = ref
    log.info("Merge cells: ".. ref)
  end
  self.merge_cells = merge_cells
end

function Sheet:save_dimensions(dom)
  -- parse the dimension info to find number of columns and rows
  local dimobj = dom:query_selector("dimension")
  if #dimobj > 0 then
    local dimension = dimobj[1]:get_attribute("ref")
    log.info("Sheet dimensions: " .. dimension)
    -- what is the number of columns and rows? we should construct the table where
    -- each row has number of columns equal to self.columns, in order to support the
    -- ranges properly
    self.left,self.top,self.columns,self.rows = ranges.get_range(dimension)
    return true
  else
    log.warn("Cannot find table dimensions")
    return false
  end
end

-- named ranges enables us to reference to cells using names, rather than A1:B2 ranges
-- but the current form doesn't work, it seems that named ranges are saved in another file,
-- in xl/tables/ directory
function Sheet:load_named_ranges(dom)
  local table_parts = dom:query_selector("tablePart")
  local named_ranges = {}
  for _, range in ipairs(table_parts) do
    local rid = range:get_attribute("r:id")
    local filename = self:find_file_by_id(rid, self.directory)
    if filename then
      local range_table = self:load_zip_xml(filename)
      local tbl = range_table:query_selector("table")[1]
      local name = tbl:get_attribute("displayname")
      local ref = tbl:get_attribute("ref")
      named_ranges[name] = ref
      log.info("Named range " .. name .. ": ".. ref)
    end
    -- local name = range:get_attribute("name")
    -- local content = range:get_text()
    -- named_ranges[name] = content
  end
  self.named_ranges = named_ranges
end

function Sheet:process_rows(dom)
  local rows = {}
  lastn = 0
  for _, row in ipairs(dom:query_selector("row")) do
    -- there may be empty rows, we must add blank rows to the generated table
    local n = tonumber(row:get_attribute("r"))
    if n > self.rows then break end
    -- if the diff is bigger than 1, there are empty rows
    -- what to do with them? the renderer should handle them.
    local diff = n - lastn
    -- prepare table with empty columns according to table column count
    local column = self:prepare_row()
    -- the
    column.index = n
    log.info("Row: ".. n.. " width "..#column)
    for _, cell in ipairs(row:query_selector("c")) do
      local pos, content = self:parse_cell(cell)
      column[pos]= content
    end
    rows[n] = column
    lastn = n
  end
  self.table = rows

  -- xxx("sheetViews")
  -- log.info("sheet ".. name .. " has " ..rows .. " rows")
end

function Sheet:prepare_row()
  local t = {}
  -- make table with empty columns according to table width
  -- it must be number indexed table with value field with empty string
  for i=1, self.columns do
    t[i] = {{value=""}}
  end
  return t
end

--- Parse cell contents
function Sheet:parse_cell(cell)
  local range = cell:get_attribute("r")
  local style = cell:get_attribute("s")
  local t = cell:get_attribute("t")
  -- find horizontal position in the row
  -- it suffices to get only the first dimension
  local pos = ranges.get_range(range)
  -- get the shared strings, replace the current value of cell
  if t == "s" then
    local ref = tonumber(cell:get_text())
    cell = self.shared_strings[ref]
  end
  local value = self:handle_values(cell)
  value.style = self:get_cell_style(style)
  -- handle an empty cell
  return pos, value
end



function Sheet:handle_values(cell)
  local children = {}
  local current_style = {}
  local name = cell:get_element_name()
  log.info("top cell: ".. name)
  -- the cell can contain values in v element, or inline strings in t elements
  -- we don't need to handle inline styles
  if name == "c" then
    local elements = cell:query_selector("v,t")
    for _, el in ipairs(elements) do
      local value = el:get_text()
      table.insert(children, {value = value})
    end
  -- the text from shared text table. it may contain rich text styles, which should be handled
  else
    for _, el in ipairs(cell:query_selector("t")) do
      local value = el:get_text()
      local prev = el:get_prev_node()
      -- is rPr element always placed before the t element? In the specification it always is.
      if prev and prev:is_element() and prev:get_element_name() == "rPr" then
        current_style = self:get_inline_style(prev)
      end
      table.insert(children, {value = value, style = current_style})
      -- reset the current style
      current_style = {}
    end
  end

  for _, k in ipairs(children) do
    log.info("element text: " .. string.format('"%s"',k.value))
    if k.style then
      for x,y in pairs(k.style) do
        log.info("inline style: ".. x .. " : ".. tostring(y))
      end
    end
  end
  return children
end

function Sheet:get_inline_style(el)
  local style = {}
  el:traverse_elements(function(curr)
    local name = curr:get_element_name()
    if name == "b" then
      style.bold = true
    elseif name == "i" then
      style.italic = true
    elseif name == "sz" then
      style.size = curr:get_attribute("val")
    elseif name == "color" then
      -- we support only rgb color, not the indexed palletes
      style.color = curr:get_attribute("rgb")
    end
  end)
  return style
end

function Sheet:get_cell_style(style)
  log.info("get cell style", style)
end
function Sheet:load_columns(dom)
  local columns = dom:query_selector("cols")[1]
  -- print(columns:serialize())

end

function Sheet:load_links(dom, filename)
  local links = {}
  for _, link in ipairs(dom:query_selector("hyperlink")) do
    local ref = link:get_attribute("ref")
    -- display is text in the cell which forms the hyperlink
    local display = link:get_attribute("display")
    local rid = link:get_attribute("r:id")
    local href = self:find_file_by_id(rid, filename)
    log.info("link: ".. ref .. " : " .. (href or ""))
    -- there can be only one hyperlink in one cell, so we can use the 
    -- ref as table key for fast acces.
    links[ref]= {link = href, display = display}
  end
  self.links = links
end

--- Save the loaded links to cells where they were used
function Sheet:save_links()
  local links = self.links
  local tbl = self.table
  for ref, data in pairs(links) do
    local x,y = ranges.get_range(ref)
    local cell = tbl[y][x]
    if not cell then
      log.error("Cannot apply link " .. data.link .. " to cell ".. ref)
    else
      -- find linked text in cell text parts
      if #cell > 1 then 
        for _,v in ipairs(cell) do
          if v.value == data.display then
            local style = v.style or {}
            style.link = data.link
            v.style = style
          end
        end
      elseif #cell == 1 then
        -- there is just one text in the cell
        local v = cell[1]
        local style = v.style or {}
        style.link = data.link or data.display
        v.style = style

      end
    end
  end
end
M.load = load
return M
