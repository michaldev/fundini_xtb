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

  local sheet, sheet_err = ctx.api.parse_xlsx("Cash Operations")
  if not sheet or not sheet.rows then
    ctx.log("cannot parse xlsx: " .. tostring(sheet_err))
    return { transactions = {} }
  end

  local rows = sheet.rows
  ctx.log("sheet: " .. tostring(sheet.sheet) .. ", rows: " .. tostring(sheet.row_count))

  local header_row = nil
  local cols = nil

  for i = 1, math.min(20, #rows) do
    local r = rows[i]
    if r then
      local map = {}
      for j = 1, #r do
        local key = trim(r[j])
        if key and key ~= "" then
          key = key:lower()
          if map[key] == nil then
            map[key] = j
          end
        end
      end
      if map["type"] and map["time"] and map["amount"] then
        header_row = i
        cols = map
        break
      end
    end
  end

  if not header_row then
    ctx.log("Header row not found")
    return { transactions = {} }
  end

  ctx.log("header row: " .. tostring(header_row)
    .. ", type=" .. tostring(cols["type"])
    .. ", ticker=" .. tostring(cols["ticker"])
    .. ", time=" .. tostring(cols["time"])
    .. ", amount=" .. tostring(cols["amount"])
    .. ", id=" .. tostring(cols["id"])
    .. ", comment=" .. tostring(cols["comment"])
    .. ", product=" .. tostring(cols["product"]))

  local function cell(r, name)
    local idx = cols[name]
    if idx == nil then return nil end
    return r[idx]
  end

  local transactions = {}
  local cash_operations = {}
  local skipped = 0

  for i = header_row + 1, #rows do
    local r = rows[i]

    local typ = trim(cell(r, "type"))
    local ticker = trim(cell(r, "ticker"))
    local time_serial = cell(r, "time")
    local amount = parse_number(cell(r, "amount"))
    local comment = cell(r, "comment")
    local product = trim(cell(r, "product"))
    local op_id = trim(cell(r, "id"))

    if typ == nil or typ == "" or typ == "Total" then
      goto continue
    end
    if amount == nil then
      skipped = skipped + 1
      goto continue
    end

    local time_iso, err = ctx.api.parse_excel_date(time_serial)
    if not time_iso then
      ctx.log("cannot parse date: " .. tostring(err))
      skipped = skipped + 1
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

      if units and price and units > 0 and price > 0 and amount < 0 then
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
      else
        skipped = skipped + 1
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

      if units and price and units > 0 and price > 0 and amount > 0 then
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
      else
        skipped = skipped + 1
      end

    elseif typ == "IKE deposit"
        or typ == "IKE cash transfer in"
        or typ == "IKZE deposit"
        or typ == "Deposit"
        or typ == "Cash transfer in"
        or typ == "Withdrawal"
        or typ == "IKE cash transfer out"
        or typ == "IKZE withdrawal"
        or typ == "Cash transfer out"
    then
      -- XTB names both legs of an internal transfer after the receiving account
      -- (e.g. "IKE deposit" appears with a negative amount on the source account),
      -- so the direction comes from the sign, not from the label.
      local kind = "deposit"
      if amount < 0 then
        kind = "withdrawal"
      end
      table.insert(cash_operations, {
        date = time_iso,
        type = kind,
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

    else
      skipped = skipped + 1
      ctx.log("unhandled type: " .. typ)
    end

    ::continue::
  end

  ctx.log("Transactions created: " .. tostring(#transactions))
  ctx.log("Cash operations created: " .. tostring(#cash_operations))
  ctx.log("Rows skipped: " .. tostring(skipped))
  return {
    transactions = transactions,
    cash_operations = cash_operations,
  }
end
