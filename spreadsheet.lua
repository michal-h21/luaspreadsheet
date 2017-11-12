kpse.set_program_name "luatex"
local xslt = require "spreadsheet.xslt-reader"

local obj,msg  = xslt.load("pokus.xlsx")

if not obj then
  print(msg)
  os.exit()
end


