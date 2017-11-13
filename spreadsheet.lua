kpse.set_program_name "luatex"
local xslt = require "spreadsheet.xslt-reader"

-- local obj,msg  = xslt.load("pokus.xlsx")
local obj,msg  = xslt.load("odpis.xlsx")

if not obj then
  print(msg)
  os.exit()
end

local sheet = obj:get_sheet("Sheet1")
local sheet = obj:get_sheet("List1")
local sheet = obj:get_sheet(1)

