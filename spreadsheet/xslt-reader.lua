--- XSLT reader module
local M = {}
local zip = require "zip"
local domobject = require "luaxml-domobject"
local log = require "spreadsheet.log"


--- @type Xlsx object
local Xlsx = {}

--- load the xlsx file
-- @return Xlsx object
local function load(filename)
  -- to add support for --reader option of TeX, we may use the following trick
  local f = io.open(filename, "r")
  if not f then return nil, "File ".. filename .. " not found" end
  f:close()
  local xlsx_obj = setmetatable({}, Xlsx)
  Xlsx.__index = Xlsx
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
  return load_zip_xml(zip_file,filename)
end

function Xlsx:load_workbook(filename)
  local workbook,msg = self:load_zip_xml(filename)
  -- get filenames of particular worksheets
  local directory = filename:match("(.-)[^/]+$")
  workbook.directory = directory
  log.info("workbook path:".. directory)
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

--- Retrieve sheet from workbook.
-- It can be referenced either by name, or id number.
function Xlsx:get_sheet(
  name -- string or number
)
  local name = name or 1
  local workbook = self.workbook
  local attr = "name"
  -- the path to sheet file is saved in the relationships table
  -- there are several of such tables, for different directories
  -- we must retrieve the one for the directory where the workbook
  -- file lies
  local relationships = self.relationships or {}
  local directory = workbook.directory
  local rel_table = relationships[directory] or {}
  if type(name) == "number" then
    attr = "sheetid"
    -- the attribute value is string, so we must convert the name
    -- to string to assure they match
    name = tostring(name)
  end
  -- print(self.workbook:serialize())
  local selected
  for _, sheet in ipairs(workbook:query_selector("sheets sheet")) do
    if sheet:get_attribute(attr) == name then
      -- selected = 
      local rid = sheet:get_attribute("r:id")
      local relation_dest = rel_table[rid]
      if relation_dest then
        local target = relation_dest.target
        log.info("found sheet", sheet:get_attribute("name"), rid, target)
        return target
      else
        local msg = "Cannot find sheet " .. name
        log.error(msg)
        return nil, msg
      end
    end
  end
end

M.load = load
return M
