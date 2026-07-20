local function parse_number(v)
  if v == nil then return nil end
  local s = tostring(v):gsub("%s+", ""):gsub(",", ".")
  if s == "" then return nil end
  return tonumber(s)
end

local function trim(s)
  if s == nil then return nil end
  return tostring(s):match("^%s*(.-)%s*$")
end

function run(ctx)
  ctx.log("XTB Cash Operations importer start")

  local sheet = ctx.api.parse_xlsx()
  if not sheet or not sheet.rows then
    return { transactions = {} }
  end

  local rows = sheet.rows

  local header_row = nil
  for i = 1, math.min(20, #rows) do
    local r = rows[i]
    if r and trim(r[1]) == "Type" and trim(r[2]) == "Ticker" then
      header_row = i
      break
    end
  end

  if not header_row then
    ctx.log("Header row not found")
    return { transactions = {} }
  end

  local transactions = {}
  local cash_operations = {}

  for i = header_row + 1, #rows do
    local r = rows[i]

    local typ = trim(r[1])
    local ticker = trim(r[2])
    local time_serial = r[4]
    local amount = parse_number(r[5])
    local comment = r[7]
    local product = trim(r[8])
    local op_id = trim(r[6])

    if typ == nil or typ == "" or typ == "Total" then
      goto continue
    end
    if amount == nil then
      goto continue
    end

    local time_iso, err = ctx.api.parse_excel_date(time_serial)
    if not time_iso then
      ctx.log("cannot parse date: " .. tostring(err))
      goto continue
    end

    local ticker_ptr = nil
    if ticker and ticker ~= "" then
      ticker_ptr = ticker
    end

    if typ == "Stock purchase" then
      local units, price
      if comment then
        -- XTB partial fills use "OPEN BUY <units>/<order_total> @ <price>";
        -- capture the per-fill units (numerator), ignoring the optional "/total".
        units, price = comment:match("OPEN BUY ([%d%.]+)/?[%d%.]* @ ([%d%.]+)")
      end
      units = tonumber(units)
      price = tonumber(price)

      if units and price and amount < 0 then
        local total_portfolio = math.abs(amount)
        local price_portfolio = total_portfolio / units
        table.insert(transactions, {
          ticker = ticker,
          trade_datetime = time_iso,
          side = "buy",
          units = units,
          instrument_currency = nil,
          price_instrument = price,
          fx_rate = price_portfolio / price,
          price_portfolio = price_portfolio,
          total_portfolio = total_portfolio,
          fee_portfolio = 0,
          tax_portfolio = 0,
          note = "XTB cash operation",
          import_name = ticker,
        })
      end

    elseif typ == "Stock sell" then
      local units, price
      if comment then
        -- XTB partial fills use "CLOSE BUY <units>/<order_total> @ <price>";
        -- capture the per-fill units (numerator), ignoring the optional "/total".
        units, price = comment:match("CLOSE BUY ([%d%.]+)/?[%d%.]* @ ([%d%.]+)")
      end
      units = tonumber(units)
      price = tonumber(price)

      if units and price and amount > 0 then
        local total_portfolio = amount
        local price_portfolio = total_portfolio / units
        table.insert(transactions, {
          ticker = ticker,
          trade_datetime = time_iso,
          side = "sell",
          units = units,
          instrument_currency = nil,
          price_instrument = price,
          fx_rate = price_portfolio / price,
          price_portfolio = price_portfolio,
          total_portfolio = total_portfolio,
          fee_portfolio = 0,
          tax_portfolio = 0,
          note = "XTB cash operation",
          import_name = ticker,
        })
      end

    elseif typ == "IKE deposit"
        or typ == "IKE cash transfer in"
        or typ == "IKZE deposit"
        or typ == "Deposit"
        or typ == "Cash transfer in"
    then
      table.insert(cash_operations, {
        date = time_iso,
        type = "deposit",
        amount_portfolio = amount,
        ticker = nil,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Withdrawal"
        or typ == "IKE cash transfer out"
        or typ == "IKZE withdrawal"
        or typ == "Cash transfer out"
    then
      table.insert(cash_operations, {
        date = time_iso,
        type = "withdrawal",
        amount_portfolio = amount,
        ticker = nil,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Dividend" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "dividend",
        amount_portfolio = amount,
        ticker = ticker_ptr,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Withholding tax" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "tax",
        amount_portfolio = amount,
        ticker = ticker_ptr,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "SEC fee" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "fee",
        amount_portfolio = amount,
        ticker = ticker_ptr,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Free funds interest" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "interest",
        amount_portfolio = amount,
        ticker = nil,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Correction" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "adjustment",
        amount_portfolio = amount,
        ticker = ticker_ptr,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Close trade" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "adjustment",
        amount_portfolio = amount,
        ticker = ticker_ptr,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Fractional shares" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "adjustment",
        amount_portfolio = amount,
        ticker = ticker_ptr,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Swap" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "fee",
        amount_portfolio = amount,
        ticker = ticker_ptr,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Commission" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "fee",
        amount_portfolio = amount,
        ticker = ticker_ptr,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Free funds interest tax" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "tax",
        amount_portfolio = amount,
        ticker = nil,
        note = comment,
        import_name = product,
        external_id = op_id,
      })

    elseif typ == "Transfer" then
      table.insert(cash_operations, {
        date = time_iso,
        type = "transfer",
        amount_portfolio = amount,
        ticker = nil,
        note = comment,
        import_name = product,
        external_id = op_id,
      })
    end

    ::continue::
  end

  ctx.log("Transactions created: " .. tostring(#transactions))
  ctx.log("Cash operations created: " .. tostring(#cash_operations))
  return {
    transactions = transactions,
    cash_operations = cash_operations,
  }
end
