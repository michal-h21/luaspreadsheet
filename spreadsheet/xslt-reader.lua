--- XSLT reader module
local M = {}
local zip = require "zip"
local domobject = require "luaxml-domobject"
local log = require "spreadsheet.log"
local ranges = require "spreadsheet.ranges"


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
  self.log = {}
  -- load file with pointer to other files
  local content_types, msg = self:load_zip_xml("[Content_Types].xml")
  if not content_types then
    return nil, msg
  end
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
  if not dom then 
    log.error("Cannot parse XML from " .. filename)
  end
  return dom, msg
end

function Xlsx:load_workbook(filename)
  local workbook,msg = self:load_zip_xml(filename)
  -- because the paths in _rels table are relative to the current 
  -- directory, we must save the current path, which will be then 
  -- used in the file loading
  local directory = filename:match("(.-)[^/]+$")
  workbook.directory = directory
  log.info("workbook path:".. directory)
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

function Xlsx:load_relationships(filename)
  -- each relationship file correspond to the parent directory
  -- so we must find that corresponding dir
  local path = filename:match("(.-)_rels")
  local relationships = self.relationships or {}
  local current = {}
  local dom = self:load_zip_xml(filename)
  for _, el in ipairs(dom:query_selector("Relationship")) do
    local id, target, schema = el:get_attribute("id"), el:get_attribute("target"), el:get_attribute("type") 
    -- we must construct full path to the target file
    local fulltarget = path .. target
    current[id] = {target = fulltarget, type = schema}
  end
  relationships[path] = current
  self.relationships = relationships
end

function Xlsx:load_styles(filename)
  self.styles = self:load_zip_xml(filename)
end

function Xlsx:load_shared_strings(filename)
  self.shared_strins = self:load_zip_xml(filename)
end

function Xlsx:find_file_by_id(rid)
  local workbook = self.workbook
  -- the path to sheet file is saved in the relationships table
  -- there are several of such tables, for different directories
  -- we must retrieve the one for the directory where the workbook
  -- file lies
  local relationships = self.relationships or {}
  local directory = workbook.directory
  local rel_table = relationships[directory] or {}
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
  local dom = self:load_zip_xml(name)
  local sheet_obj = setmetatable({}, Sheet)
  sheet_obj:set_parent(self)
  sheet_obj:load_dom(name, dom)
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

function Sheet:set_parent(parent)
  self.parent = parent
end

function Sheet:load_dom(name, dom)
  local rows = #dom:query_selector("row")
  local function xxx(sel)
    print(sel,dom:query_selector(sel)[1]:serialize())
  end
  -- xxx("sheetData")
  local dimension = dom:query_selector("dimension")[1]:get_attribute("ref")
  print(ranges.get_range(dimension))
  -- xxx("dimension")
  xxx("sheetViews")
  xxx("cols")
  log.info("sheet ".. name .. " has " ..rows .. " rows")
  return dom
end

M.load = load
return M
