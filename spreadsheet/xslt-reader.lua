--- XSLT reader module
local M = {}
local zip = require "zip"
local domobject = require "luaxml-domobject"


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
  -- load file with pointer to other files
  local content_types, msg = self:load_zip_xml("[Content_Types].xml")
  if not content_types then
    return nil, msg
  end
  self.content_types = content_types
  for _, override in ipairs(content_types:query_selector("Override")) do
    for k,v in pairs(override._attr) do
      print(k,v)
    end
    -- print(override:get_attribute("PartName"), override:get_attribute("ContentType"))
  end
  return true
end

function Xlsx:load_zip_xml(filename)
  local zip_file = self.file
  return load_zip_xml(zip_file,filename)
end


M.load = load
return M
