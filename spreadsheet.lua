kpse.set_program_name "luatex"
local xlsx = require "spreadsheet.xlsx-reader"

local lo,msg  = xlsx.load("pokus.xlsx")

-- local obj,msg  = xslt.load("odpis.xlsx")

if not obj and not lo then
  print(msg)
  os.exit()
end

local sheet = lo:get_sheet("Sheet1")
-- local sheet = obj:get_sheet("Sheet1")
-- local sheet = obj:get_sheet("List1")
-- local sheet = obj:get_sheet(1)

