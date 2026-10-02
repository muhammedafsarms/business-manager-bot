import os
from datetime import datetime
from openpyxl import Workbook, load_workbook
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# =========================
# BOT TOKEN
# =========================

BOT_TOKEN = os.environ.get("BOT_TOKEN")

# =========================
# EXCEL
# =========================

EXCEL_FILE = os.environ.get("EXCEL_FILE", "business.xlsx")


def create_excel():
    if os.path.exists(EXCEL_FILE):
        return

    wb = Workbook()

    ws = wb.active
    ws.title = "Products"

    ws.append([
        "ID",
        "Product",
        "Buy Price",
        "Sell Price",
        "Stock"
    ])

    sales = wb.create_sheet("Sales")
    sales.append([
        "ID",
        "Product",
        "Quantity",
        "Price",
        "Total",
        "Customer",
        "Date",
        "Time"
    ])

    expenses = wb.create_sheet("Expenses")
    expenses.append([
        "ID",
        "Description",
        "Amount",
        "Date",
        "Time"
    ])

    customers = wb.create_sheet("Customers")
    customers.append([
        "ID",
        "Name",
        "Phone",
        "Notes"
    ])

    wb.save(EXCEL_FILE)


# =========================
# MENUS
# =========================

def main_menu():
    keyboard = [
        ["📦 Products", "🛒 New Sale"],
        ["📊 Stock", "💸 Expenses"],
        ["📈 Reports", "🔍 Search"],
        ["👥 Customers", "📁 Excel"],
        ["⚙️ Settings"],
    ]

    return ReplyKeyboardMarkup(
        keyboard,
        resize_keyboard=True
    )


def products_menu():
    keyboard = [
        ["➕ Add Product", "📋 View Products"],
        ["✏️ Edit Product", "🗑 Delete Product"],
        ["🔙 Back"],
    ]

    return ReplyKeyboardMarkup(
        keyboard,
        resize_keyboard=True
    )


# =========================
# START
# =========================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    create_excel()

    context.user_data.clear()

    await update.message.reply_text(
        "🏪 *Business Manager Bot*\n\n"
        "Welcome! Your business management system is ready.\n\n"
        "Choose an option below 👇",
        reply_markup=main_menu(),
        parse_mode="Markdown"
    )


# =========================
# PRODUCTS MENU
# =========================

async def show_products(update: Update, context: ContextTypes.DEFAULT_TYPE):

    context.user_data.clear()

    await update.message.reply_text(
        "📦 *Products*\n\n"
        "Choose an option:",
        reply_markup=products_menu(),
        parse_mode="Markdown"
    )


# =========================
# ADD PRODUCT
# =========================

async def start_add_product(update: Update, context: ContextTypes.DEFAULT_TYPE):

    context.user_data.clear()
    context.user_data["state"] = "product_name"

    await update.message.reply_text(
        "➕ *Add Product*\n\n"
        "Enter the product name:",
        parse_mode="Markdown"
    )


async def add_product_steps(update: Update, context: ContextTypes.DEFAULT_TYPE):

    state = context.user_data.get("state")

    if not state:
        return

    text = update.message.text.strip()

    if state == "product_name":

        context.user_data["product_name"] = text
        context.user_data["state"] = "buy_price"

        await update.message.reply_text(
            "💰 Enter the buying price:"
        )

    elif state == "buy_price":

        try:
            price = float(text)
        except ValueError:
            await update.message.reply_text(
                "❌ Please enter a valid number."
            )
            return

        context.user_data["buy_price"] = price
        context.user_data["state"] = "sell_price"

        await update.message.reply_text(
            "🏷️ Enter the selling price:"
        )

    elif state == "sell_price":

        try:
            price = float(text)
        except ValueError:
            await update.message.reply_text(
                "❌ Please enter a valid number."
            )
            return

        context.user_data["sell_price"] = price
        context.user_data["state"] = "stock"

        await update.message.reply_text(
            "📊 Enter the stock quantity:"
        )

    elif state == "stock":

        try:
            stock = int(text)
        except ValueError:
            await update.message.reply_text(
                "❌ Please enter a whole number."
            )
            return

        context.user_data["stock"] = stock

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Products"]

        # Automatic ID
        next_id = ws.max_row

        ws.append([
            next_id,
            context.user_data["product_name"],
            context.user_data["buy_price"],
            context.user_data["sell_price"],
            stock
        ])

        wb.save(EXCEL_FILE)

        product_name = context.user_data["product_name"]
        buy_price = context.user_data["buy_price"]
        sell_price = context.user_data["sell_price"]

        context.user_data.clear()

        await update.message.reply_text(
            "✅ *Product Added Successfully!*\n\n"
            f"🆔 ID: `{next_id}`\n"
            f"📦 Product: {product_name}\n"
            f"💰 Buy Price: ₹{buy_price:g}\n"
            f"🏷️ Sell Price: ₹{sell_price:g}\n"
            f"📊 Stock: {stock}",
            reply_markup=products_menu(),
            parse_mode="Markdown"
        )


# =========================
# VIEW PRODUCTS
# =========================

async def view_products(update: Update, context: ContextTypes.DEFAULT_TYPE):

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Products"]

    if ws.max_row <= 1:

        await update.message.reply_text(
            "📦 *Products*\n\n"
            "No products added yet.",
            reply_markup=products_menu(),
            parse_mode="Markdown"
        )
        return

    message = "📦 *Your Products*\n\n"

    for row in ws.iter_rows(min_row=2, values_only=True):

        product_id, name, buy, sell, stock = row

        message += (
            f"🆔 *{product_id}* — {name}\n"
            f"💰 Buy: ₹{buy:g}\n"
            f"🏷️ Sell: ₹{sell:g}\n"
            f"📊 Stock: {stock}\n\n"
        )

    await update.message.reply_text(
        message,
        reply_markup=products_menu(),
        parse_mode="Markdown"
    )


# =========================
# DELETE PRODUCT
# =========================

async def start_delete_product(update: Update, context: ContextTypes.DEFAULT_TYPE):

    context.user_data.clear()
    context.user_data["state"] = "delete_product"

    await update.message.reply_text(
        "🗑 *Delete Product*\n\n"
        "Enter the product ID:",
        parse_mode="Markdown"
    )


async def delete_product(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if context.user_data.get("state") != "delete_product":
        return

    try:
        product_id = int(update.message.text.strip())
    except ValueError:
        await update.message.reply_text(
            "❌ Please enter a valid product ID."
        )
        return

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Products"]

    found = False
    product_name = ""

    for row in range(2, ws.max_row + 1):

        if ws.cell(row, 1).value == product_id:

            product_name = ws.cell(row, 2).value
            ws.delete_rows(row, 1)
            found = True
            break

    if not found:

        await update.message.reply_text(
            "❌ Product not found."
        )
        return

    wb.save(EXCEL_FILE)
    context.user_data.clear()

    await update.message.reply_text(
        f"🗑️ *Product Deleted*\n\n"
        f"📦 {product_name}",
        reply_markup=products_menu(),
        parse_mode="Markdown"
    )


# =========================
# EDIT PRODUCT
# =========================

async def start_edit_product(update: Update, context: ContextTypes.DEFAULT_TYPE):

    context.user_data.clear()
    context.user_data["state"] = "edit_id"

    await update.message.reply_text(
        "✏️ *Edit Product*\n\n"
        "Enter the product ID:",
        parse_mode="Markdown"
    )


async def edit_product(update: Update, context: ContextTypes.DEFAULT_TYPE):

    state = context.user_data.get("state")
    text = update.message.text.strip()

    if state == "edit_id":

        try:
            product_id = int(text)
        except ValueError:
            await update.message.reply_text(
                "❌ Enter a valid product ID."
            )
            return

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Products"]

        found = False

        for row in range(2, ws.max_row + 1):

            if ws.cell(row, 1).value == product_id:

                context.user_data["edit_row"] = row
                found = True
                break

        if not found:

            await update.message.reply_text(
                "❌ Product not found."
            )
            return

        context.user_data["state"] = "edit_name"

        await update.message.reply_text(
            "📦 Enter the new product name:"
        )

    elif state == "edit_name":

        row = context.user_data["edit_row"]

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Products"]

        ws.cell(row, 2).value = text
        wb.save(EXCEL_FILE)

        context.user_data.clear()

        await update.message.reply_text(
            "✅ Product updated successfully!",
            reply_markup=products_menu()
        )



# =========================
# NEW SALE - CLICKABLE
# =========================

async def start_new_sale(update, context):
    context.user_data.clear()

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Products"]

    buttons = []

    for row in range(2, ws.max_row + 1):
        product_id = ws.cell(row, 1).value
        name = ws.cell(row, 2).value
        stock = ws.cell(row, 5).value

        if name and stock and stock > 0:
            buttons.append([
                InlineKeyboardButton(
                    f"📦 {name} ({stock})",
                    callback_data=f"sale_product:{product_id}"
                )
            ])

    buttons.append([
        InlineKeyboardButton("🔙 Back", callback_data="sale_back")
    ])

    if len(buttons) == 1:
        await update.message.reply_text(
            "🛒 *New Sale*\n\n❌ No products are currently in stock.",
            parse_mode="Markdown"
        )
        return

    await update.message.reply_text(
        "🛒 *Select Product*",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode="Markdown"
    )


async def sale_callback(update, context):
    query = update.callback_query
    await query.answer()

    data = query.data

    if data == "sale_back":
        context.user_data.clear()

        await query.edit_message_text(
            "🏪 Main menu opened.\n\nUse the buttons below to continue."
        )
        await query.message.reply_text(
            "Choose an option:",
            reply_markup=main_menu()
        )
        return

    if data.startswith("sale_product:"):
        product_id = int(data.split(":")[1])

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Products"]

        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == product_id:

                name = ws.cell(row, 2).value
                buy_price = ws.cell(row, 3).value
                sell_price = ws.cell(row, 4).value
                stock = ws.cell(row, 5).value

                context.user_data["sale_product_id"] = product_id
                context.user_data["sale_row"] = row
                context.user_data["sale_name"] = name
                context.user_data["sale_buy"] = buy_price
                context.user_data["sale_sell"] = sell_price
                context.user_data["sale_stock"] = stock

                quantity_buttons = []

                max_quantity = min(stock, 10)

                for i in range(1, max_quantity + 1):
                    quantity_buttons.append(
                        InlineKeyboardButton(
                            str(i),
                            callback_data=f"sale_qty:{i}"
                        )
                    )

                rows = [
                    quantity_buttons[i:i + 5]
                    for i in range(0, len(quantity_buttons), 5)
                ]

                rows.append([
                    InlineKeyboardButton(
                        "🔙 Products",
                        callback_data="sale_products"
                    )
                ])

                await query.edit_message_text(
                    f"📦 *{name}*\n\n"
                    f"🏷️ Price: ₹{sell_price:g}\n"
                    f"📊 Available stock: {stock}\n\n"
                    "🔢 *Select Quantity:*",
                    reply_markup=InlineKeyboardMarkup(rows),
                    parse_mode="Markdown"
                )
                return

        await query.edit_message_text("❌ Product not found.")


    if data == "sale_products":
        await start_new_sale_callback(query, context)
        return


    if data.startswith("sale_qty:"):
        quantity = int(data.split(":")[1])

        stock = context.user_data.get("sale_stock", 0)

        if quantity > stock:
            await query.answer(
                "❌ Not enough stock!",
                show_alert=True
            )
            return

        context.user_data["sale_quantity"] = quantity

        # Load customers
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Customers"]

        buttons = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue

            customer_id, name, phone, notes = row

            buttons.append([
                InlineKeyboardButton(
                    f"👤 {name}",
                    callback_data=f"sale_customer:{customer_id}"
                )
            ])

        # Walk-in customer option
        buttons.append([
            InlineKeyboardButton(
                "🚶 Walk-in / No Customer",
                callback_data="sale_customer:0"
            )
        ])

        buttons.append([
            InlineKeyboardButton(
                "🔙 Change Quantity",
                callback_data=f"sale_product:{context.user_data['sale_product_id']}"
            )
        ])

        await query.edit_message_text(
            f"👥 *Select Customer*\n\n"
            f"📦 {context.user_data['sale_name']}\n"
            f"🔢 Quantity: {quantity}\n\n"
            "Choose a customer for this sale:",
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode="Markdown"
        )
        return


    # CUSTOMER SELECTED
    if data.startswith("sale_customer:"):

        customer_id = int(data.split(":")[1])

        context.user_data["sale_customer_id"] = customer_id

        if customer_id == 0:
            customer_name = "Walk-in Customer"

        else:
            wb = load_workbook(EXCEL_FILE, data_only=True)
            ws = wb["Customers"]

            customer_name = None

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue

                if row[0] == customer_id:
                    customer_name = row[1]
                    break

            if not customer_name:
                await query.answer(
                    "❌ Customer not found.",
                    show_alert=True
                )
                return

        context.user_data["sale_customer_name"] = customer_name

        quantity = context.user_data["sale_quantity"]
        name = context.user_data["sale_name"]
        buy_price = context.user_data["sale_buy"]
        sell_price = context.user_data["sale_sell"]

        total = sell_price * quantity
        profit = (sell_price - buy_price) * quantity

        buttons = [
            [
                InlineKeyboardButton(
                    "✅ Confirm Sale",
                    callback_data="sale_confirm"
                )
            ],
            [
                InlineKeyboardButton(
                    "❌ Cancel",
                    callback_data="sale_cancel"
                )
            ]
        ]

        await query.edit_message_text(
            "🧾 *SALE SUMMARY*\n\n"
            f"👤 Customer: {customer_name}\n"
            f"📦 Product: {name}\n"
            f"🔢 Quantity: {quantity}\n\n"
            f"🏷️ Price: ₹{sell_price:g}\n"
            f"💰 Total: ₹{total:g}\n"
            f"💵 Profit: ₹{profit:g}\n\n"
            "Confirm this sale?",
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode="Markdown"
        )
        return


    if data == "sale_cancel":
        context.user_data.clear()

        await query.edit_message_text(
            "❌ Sale cancelled."
        )
        return


    if data == "sale_confirm":

        row = context.user_data["sale_row"]
        product_name = context.user_data["sale_name"]
        buy_price = context.user_data["sale_buy"]
        sell_price = context.user_data["sale_sell"]
        quantity = context.user_data["sale_quantity"]
        stock = context.user_data["sale_stock"]

        if quantity > stock:
            await query.answer(
                "❌ Stock changed. Not enough stock.",
                show_alert=True
            )
            return

        new_stock = stock - quantity

        revenue = sell_price * quantity
        profit = (sell_price - buy_price) * quantity

        wb = load_workbook(EXCEL_FILE)

        products = wb["Products"]
        sales = wb["Sales"]

        products.cell(row, 5).value = new_stock

        sale_id = sales.max_row

        # Add Customer column to Sales sheet
        if sales.cell(1, 6).value != "Customer":
            sales.cell(1, 6).value = "Customer"

        customer_name = context.user_data.get(
            "sale_customer_name",
            "Walk-in Customer"
        )

        now = datetime.now()

        sales.append([
            sale_id,
            product_name,
            quantity,
            sell_price,
            revenue,
            customer_name,
            now.strftime("%Y-%m-%d"),
            now.strftime("%H:%M:%S")
        ])

        wb.save(EXCEL_FILE)

        context.user_data.clear()

        await query.edit_message_text(
            "✅ *SALE COMPLETED!*\n\n"
            f"🧾 Sale ID: `{sale_id}`\n"
            f"📦 {product_name}\n"
            f"🔢 Quantity: {quantity}\n\n"
            f"💰 Total: ₹{revenue:g}\n"
            f"💵 Profit: ₹{profit:g}\n"
            f"📊 Remaining stock: {new_stock}",
            parse_mode="Markdown"
        )


async def start_new_sale_callback(query, context):
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Products"]

    buttons = []

    for row in range(2, ws.max_row + 1):
        product_id = ws.cell(row, 1).value
        name = ws.cell(row, 2).value
        stock = ws.cell(row, 5).value

        if name and stock and stock > 0:
            buttons.append([
                InlineKeyboardButton(
                    f"📦 {name} ({stock})",
                    callback_data=f"sale_product:{product_id}"
                )
            ])

    buttons.append([
        InlineKeyboardButton(
            "🔙 Back",
            callback_data="sale_back"
        )
    ])

    await query.edit_message_text(
        "🛒 *Select Product*",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode="Markdown"
    )



# =========================
# EXPENSES
# =========================

async def start_expense(update, context):
    context.user_data.clear()
    context.user_data["state"] = "expense_description"

    await update.message.reply_text(
        "💸 *Add Expense*\n\n"
        "Enter the expense description:\n\n"
        "Example: Packaging",
        parse_mode="Markdown"
    )


async def expense_steps(update, context):
    state = context.user_data.get("state")
    text = update.message.text.strip()

    if state == "expense_description":
        context.user_data["expense_description"] = text
        context.user_data["state"] = "expense_amount"

        await update.message.reply_text(
            "💰 Enter the expense amount:\n\n"
            "Example: 150"
        )
        return

    if state == "expense_amount":
        try:
            amount = float(text)
        except ValueError:
            await update.message.reply_text(
                "❌ Please enter a valid amount."
            )
            return

        if amount <= 0:
            await update.message.reply_text(
                "❌ Amount must be greater than 0."
            )
            return

        description = context.user_data["expense_description"]

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Expenses"]

        expense_id = ws.max_row

        now = datetime.now()

        ws.append([
            expense_id,
            description,
            amount,
            now.strftime("%Y-%m-%d"),
            now.strftime("%H:%M:%S")
        ])

        wb.save(EXCEL_FILE)

        context.user_data.clear()

        await update.message.reply_text(
            "✅ *Expense Added!*\n\n"
            f"🧾 ID: `{expense_id}`\n"
            f"📝 {description}\n"
            f"💸 Amount: ₹{amount:g}",
            reply_markup=main_menu(),
            parse_mode="Markdown"
        )
        return


async def view_expenses(update, context):
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Expenses"]

    if ws.max_row <= 1:
        await update.message.reply_text(
            "💸 *Expenses*\n\n"
            "No expenses recorded yet.",
            parse_mode="Markdown"
        )
        return

    message = "💸 *Recent Expenses*\n\n"

    total = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        expense_id, description, amount = row[:3]

        if amount is None:
            continue

        total += float(amount)

        message += (
            f"🧾 *{expense_id}* — {description}\n"
            f"💰 ₹{float(amount):g}\n\n"
        )

    message += f"━━━━━━━━━━━━\n💸 *Total: ₹{total:g}*"

    await update.message.reply_text(
        message,
        reply_markup=main_menu(),
        parse_mode="Markdown"
    )


# =========================
# REPORTS
# =========================

async def show_reports(update, context):

    wb = load_workbook(EXCEL_FILE, data_only=True)

    products = wb["Products"]
    sales = wb["Sales"]
    expenses = wb["Expenses"]

    total_revenue = 0
    total_cost = 0
    total_profit = 0
    total_sales = 0
    total_expenses = 0

    # Sales
    for row in sales.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        quantity = row[2] or 0
        sell_price = row[3] or 0
        revenue = row[4] or 0

        total_sales += 1
        total_revenue += float(revenue)
        total_cost += float(sell_price) * 0

    # Calculate actual product cost from Products sheet
    product_costs = {}

    for row in products.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        name = row[1]
        buy_price = row[2] or 0

        product_costs[str(name)] = float(buy_price)

    for row in sales.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        name = str(row[1])
        quantity = row[2] or 0

        buy_price = product_costs.get(name, 0)

        total_cost += float(buy_price) * int(quantity)

    total_profit = total_revenue - total_cost

    # Expenses
    for row in expenses.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        amount = row[2] or 0
        total_expenses += float(amount)

    net_profit = total_profit - total_expenses

    # Products
    product_count = 0
    total_stock_units = 0
    stock_value = 0

    for row in products.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        product_count += 1

        stock = row[4] or 0
        buy_price = row[2] or 0

        total_stock_units += int(stock)
        stock_value += float(stock) * float(buy_price)

    keyboard = [
        ["🔄 Refresh", "📋 Sales"],
        ["💸 Expenses", "📦 Stock"],
        ["🔙 Back"]
    ]

    await update.message.reply_text(
        "📈 *BUSINESS REPORT*\n\n"
        "━━━━━━━━━━━━━━\n"
        f"💰 Total Revenue: ₹{total_revenue:g}\n"
        f"🛒 Total Sales: {total_sales}\n"
        f"📦 Product Cost: ₹{total_cost:g}\n"
        f"📈 Gross Profit: ₹{total_profit:g}\n"
        "━━━━━━━━━━━━━━\n"
        f"💸 Expenses: ₹{total_expenses:g}\n"
        f"💵 Net Profit: ₹{net_profit:g}\n"
        "━━━━━━━━━━━━━━\n"
        f"📦 Products: {product_count}\n"
        f"📊 Stock Units: {total_stock_units}\n"
        f"💎 Stock Value: ₹{stock_value:g}",
        reply_markup=ReplyKeyboardMarkup(
            keyboard,
            resize_keyboard=True
        ),
        parse_mode="Markdown"
    )


async def report_sales(update, context):

    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb["Sales"]

    if ws.max_row <= 1:
        await update.message.reply_text(
            "🛒 No sales recorded yet."
        )
        return

    message = "🛒 *SALES HISTORY*\n\n"
    total = 0

    for row in ws.iter_rows(min_row=2, values_only=True):

        if not row or row[0] is None:
            continue

        sale_id, product, quantity, price, revenue = row[:5]
        customer = row[5] if len(row) > 5 else "Walk-in Customer"

        total += float(revenue or 0)

        message += (
            f"🧾 #{sale_id}\n"
            f"📦 {product}\n"
            f"🔢 Qty: {quantity}\n"
            f"💰 ₹{float(revenue or 0):g}\n\n"
        )

    message += f"━━━━━━━━━━━━\n💰 *Total: ₹{total:g}*"

    await update.message.reply_text(
        message,
        parse_mode="Markdown"
    )


async def report_stock(update, context):

    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb["Products"]

    if ws.max_row <= 1:
        await update.message.reply_text(
            "📦 No products available."
        )
        return

    message = "📦 *CURRENT STOCK*\n\n"

    for row in ws.iter_rows(min_row=2, values_only=True):

        if not row or row[0] is None:
            continue

        product_id, name, buy, sell, stock = row

        message += (
            f"🆔 {product_id} — {name}\n"
            f"📊 Stock: {stock}\n"
            f"🏷️ Price: ₹{float(sell or 0):g}\n\n"
        )

    await update.message.reply_text(
        message,
        parse_mode="Markdown"
    )


# =========================
# SEARCH
# =========================

async def start_search(update, context):
    context.user_data.clear()

    keyboard = [
        ["🔎 Search Products", "🔎 Search Sales"],
        ["🔎 Search Expenses", "🔙 Back"]
    ]

    await update.message.reply_text(
        "🔍 *Search*\n\nChoose what you want to search:",
        reply_markup=ReplyKeyboardMarkup(
            keyboard,
            resize_keyboard=True
        ),
        parse_mode="Markdown"
    )


async def search_products_start(update, context):
    context.user_data.clear()
    context.user_data["state"] = "search_products"

    await update.message.reply_text(
        "📦 *Search Products*\n\n"
        "Enter a product name:",
        parse_mode="Markdown"
    )


async def search_sales_start(update, context):
    context.user_data.clear()
    context.user_data["state"] = "search_sales"

    await update.message.reply_text(
        "🛒 *Search Sales*\n\n"
        "Enter a product name:",
        parse_mode="Markdown"
    )


async def search_expenses_start(update, context):
    context.user_data.clear()
    context.user_data["state"] = "search_expenses"

    await update.message.reply_text(
        "💸 *Search Expenses*\n\n"
        "Enter an expense description:",
        parse_mode="Markdown"
    )


async def search_steps(update, context):

    state = context.user_data.get("state")
    query = update.message.text.strip().lower()

    wb = load_workbook(EXCEL_FILE, data_only=True)

    if state == "search_products":

        ws = wb["Products"]
        results = []

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row or row[0] is None:
                continue

            product_id, name, buy, sell, stock = row

            if query in str(name).lower():

                results.append(
                    f"📦 *{name}*\n"
                    f"🆔 ID: {product_id}\n"
                    f"💰 Buy: ₹{float(buy or 0):g}\n"
                    f"🏷️ Sell: ₹{float(sell or 0):g}\n"
                    f"📊 Stock: {stock}"
                )

        if not results:
            message = "❌ No products found."
        else:
            message = "🔍 *Product Results*\n\n" + "\n\n".join(results)

        context.user_data.clear()

        await update.message.reply_text(
            message,
            reply_markup=main_menu(),
            parse_mode="Markdown"
        )

    elif state == "search_sales":

        ws = wb["Sales"]
        results = []
        total = 0

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row or row[0] is None:
                continue

            sale_id, product, quantity, price, revenue = row

            if query in str(product).lower():

                revenue = float(revenue or 0)
                total += revenue

                results.append(
                    f"🧾 Sale #{sale_id}\n"
                    f"📦 {product}\n"
                    f"🔢 Qty: {quantity}\n"
                    f"💰 ₹{revenue:g}"
                )

        if not results:
            message = "❌ No sales found."
        else:
            message = (
                "🔍 *Sales Results*\n\n"
                + "\n\n".join(results)
                + f"\n\n━━━━━━━━━━━━\n"
                f"💰 Total: ₹{total:g}"
            )

        context.user_data.clear()

        await update.message.reply_text(
            message,
            reply_markup=main_menu(),
            parse_mode="Markdown"
        )

    elif state == "search_expenses":

        ws = wb["Expenses"]
        results = []
        total = 0

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row or row[0] is None:
                continue

            expense_id, description, amount = row[:3]

            if query in str(description).lower():

                amount = float(amount or 0)
                total += amount

                results.append(
                    f"🧾 Expense #{expense_id}\n"
                    f"📝 {description}\n"
                    f"💸 ₹{amount:g}"
                )

        if not results:
            message = "❌ No expenses found."
        else:
            message = (
                "🔍 *Expense Results*\n\n"
                + "\n\n".join(results)
                + f"\n\n━━━━━━━━━━━━\n"
                f"💸 Total: ₹{total:g}"
            )

        context.user_data.clear()

        await update.message.reply_text(
            message,
            reply_markup=main_menu(),
            parse_mode="Markdown"
        )


# =========================
# CUSTOMERS
# =========================

def customer_menu():
    keyboard = [
        ["➕ Add Customer", "📋 Customer List"],
        ["🔍 Search Customer", "✏️ Edit Customer"],
        ["🗑 Delete Customer", "🔙 Back"],
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


async def show_customers(update, context):
    context.user_data.clear()

    await update.message.reply_text(
        "👥 *Customers*\n\nChoose an option:",
        reply_markup=customer_menu(),
        parse_mode="Markdown"
    )


async def start_add_customer(update, context):
    context.user_data.clear()
    context.user_data["state"] = "customer_name"

    await update.message.reply_text(
        "➕ *Add Customer*\n\nEnter customer name:",
        parse_mode="Markdown"
    )


async def customer_steps(update, context):
    state = context.user_data.get("state")
    text = update.message.text.strip()

    # ADD CUSTOMER
    if state == "customer_name":
        context.user_data["customer_name"] = text
        context.user_data["state"] = "customer_phone"

        await update.message.reply_text(
            "📱 Enter phone number:\n\n"
            "Type `skip` if you don't want to save one.",
            parse_mode="Markdown"
        )
        return

    if state == "customer_phone":
        context.user_data["customer_phone"] = (
            "" if text.lower() == "skip" else text
        )
        context.user_data["state"] = "customer_notes"

        await update.message.reply_text(
            "📝 Enter notes:\n\n"
            "Type `skip` if you don't want to add notes.",
            parse_mode="Markdown"
        )
        return

    if state == "customer_notes":
        notes = "" if text.lower() == "skip" else text

        name = context.user_data["customer_name"]
        phone = context.user_data["customer_phone"]

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Customers"]

        customer_id = ws.max_row

        ws.append([
            customer_id,
            name,
            phone,
            notes
        ])

        wb.save(EXCEL_FILE)
        context.user_data.clear()

        await update.message.reply_text(
            "✅ *Customer Added!*\n\n"
            f"🆔 ID: `{customer_id}`\n"
            f"👤 Name: {name}\n"
            f"📱 Phone: {phone or 'Not added'}\n"
            f"📝 Notes: {notes or 'None'}",
            reply_markup=customer_menu(),
            parse_mode="Markdown"
        )
        return

    # SEARCH CUSTOMER
    if state == "search_customer":
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Customers"]

        results = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue

            customer_id, name, phone, notes = row

            if (
                text.lower() in str(name or "").lower()
                or text.lower() in str(phone or "").lower()
            ):
                results.append(
                    f"👤 *{name}*\n"
                    f"🆔 ID: {customer_id}\n"
                    f"📱 Phone: {phone or 'Not added'}\n"
                    f"📝 Notes: {notes or 'None'}"
                )

        context.user_data.clear()

        message = (
            "❌ No customers found."
            if not results
            else "🔍 *Customer Results*\n\n" + "\n\n".join(results)
        )

        await update.message.reply_text(
            message,
            reply_markup=customer_menu(),
            parse_mode="Markdown"
        )
        return

    # DELETE CUSTOMER
    if state == "delete_customer":
        try:
            customer_id = int(text)
        except ValueError:
            await update.message.reply_text(
                "❌ Enter a valid customer ID."
            )
            return

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Customers"]

        found = False
        customer_name = ""

        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == customer_id:
                customer_name = ws.cell(row, 2).value
                ws.delete_rows(row, 1)
                found = True
                break

        if not found:
            await update.message.reply_text(
                "❌ Customer not found."
            )
            return

        wb.save(EXCEL_FILE)
        context.user_data.clear()

        await update.message.reply_text(
            f"🗑️ *Customer Deleted*\n\n👤 {customer_name}",
            reply_markup=customer_menu(),
            parse_mode="Markdown"
        )
        return

    # EDIT CUSTOMER
    if state == "edit_customer_id":
        try:
            customer_id = int(text)
        except ValueError:
            await update.message.reply_text(
                "❌ Enter a valid customer ID."
            )
            return

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Customers"]

        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == customer_id:
                context.user_data["edit_customer_row"] = row
                context.user_data["state"] = "edit_customer_name"

                await update.message.reply_text(
                    "👤 Enter the new customer name:"
                )
                return

        await update.message.reply_text(
            "❌ Customer not found."
        )
        return

    if state == "edit_customer_name":
        row = context.user_data["edit_customer_row"]

        wb = load_workbook(EXCEL_FILE)
        ws = wb["Customers"]

        ws.cell(row, 2).value = text
        wb.save(EXCEL_FILE)

        context.user_data.clear()

        await update.message.reply_text(
            "✅ Customer updated successfully!",
            reply_markup=customer_menu()
        )
        return


async def customer_list(update, context):

    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb["Customers"]

    buttons = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        customer_id = row[0]
        name = row[1]

        buttons.append([
            InlineKeyboardButton(
                f"👤 {name}",
                callback_data=f"customer_view:{customer_id}"
            )
        ])

    if not buttons:
        await update.message.reply_text(
            "👥 *Customer List*\n\n"
            "No customers added yet.",
            reply_markup=customer_menu(),
            parse_mode="Markdown"
        )
        return

    await update.message.reply_text(
        "👥 *Customer List*\n\n"
        "Tap a customer to view purchase history:",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode="Markdown"
    )


async def start_search_customer(update, context):
    context.user_data.clear()
    context.user_data["state"] = "search_customer"

    await update.message.reply_text(
        "🔍 *Search Customer*\n\n"
        "Enter a customer name or phone number:",
        parse_mode="Markdown"
    )


async def start_edit_customer(update, context):
    context.user_data.clear()
    context.user_data["state"] = "edit_customer_id"

    await update.message.reply_text(
        "✏️ *Edit Customer*\n\nEnter the customer ID:",
        parse_mode="Markdown"
    )


async def start_delete_customer(update, context):
    context.user_data.clear()
    context.user_data["state"] = "delete_customer"

    await update.message.reply_text(
        "🗑 *Delete Customer*\n\nEnter the customer ID:",
        parse_mode="Markdown"
    )


# =========================
# CUSTOMER PURCHASE HISTORY
# =========================

async def customer_callback(update, context):
    query = update.callback_query
    await query.answer()

    data = query.data

    if not data.startswith("customer_view:"):
        return

    customer_id = int(data.split(":")[1])

    wb = load_workbook(EXCEL_FILE, data_only=True)

    customers = wb["Customers"]
    sales = wb["Sales"]

    customer_name = None
    customer_phone = ""

    for row in customers.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        if row[0] == customer_id:
            customer_name = row[1]
            customer_phone = row[2] or ""
            break

    if not customer_name:
        await query.edit_message_text(
            "❌ Customer not found."
        )
        return

    purchases = []
    total_spent = 0

    for row in sales.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        sale_id = row[0]
        product = row[1]
        quantity = row[2]
        revenue = row[4] or 0

        # Customer is stored in column 6
        sale_customer = row[5] if len(row) >= 6 else ""

        if str(sale_customer).strip().lower() == str(customer_name).strip().lower():

            revenue = float(revenue)

            total_spent += revenue

            purchases.append(
                f"🧾 *Sale #{sale_id}*\n"
                f"📦 {product} × {quantity}\n"
                f"💰 ₹{revenue:g}"
            )

    if purchases:
        purchase_text = "\n\n".join(purchases)
    else:
        purchase_text = "No purchases recorded yet."

    message = (
        f"👤 *{customer_name}*\n\n"
        f"📱 {customer_phone or 'No phone'}\n\n"
        f"🛒 Purchases: {len(purchases)}\n"
        f"💰 Total Spent: ₹{total_spent:g}\n\n"
        f"━━━━━━━━━━━━━━\n"
        f"{purchase_text}"
    )

    buttons = [
        [
            InlineKeyboardButton(
                "🔙 Customer List",
                callback_data="customer_list"
            )
        ]
    ]

    await query.edit_message_text(
        message,
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode="Markdown"
    )


async def customer_list_callback(update, context):
    query = update.callback_query
    await query.answer()

    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb["Customers"]

    buttons = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        customer_id = row[0]
        name = row[1]

        buttons.append([
            InlineKeyboardButton(
                f"👤 {name}",
                callback_data=f"customer_view:{customer_id}"
            )
        ])

    if not buttons:
        await query.edit_message_text(
            "👥 No customers added yet."
        )
        return

    await query.edit_message_text(
        "👥 *Customer List*\n\n"
        "Tap a customer to view their purchase history:",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode="Markdown"
    )

# =========================
# BUTTON HANDLER
# =========================

async def buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):

    state = context.user_data.get("state")

    if state in [
        "customer_name",
        "customer_phone",
        "customer_notes",
        "search_customer",
        "delete_customer",
        "edit_customer_id",
        "edit_customer_name"
    ]:
        await customer_steps(update, context)
        return


    text = update.message.text

    if text == "👥 Customers":
        await show_customers(update, context)
        return

    if text == "➕ Add Customer":
        await start_add_customer(update, context)
        return

    if text == "📋 Customer List":
        await customer_list(update, context)
        return

    if text == "🔍 Search Customer":
        await start_search_customer(update, context)
        return

    if text == "✏️ Edit Customer":
        await start_edit_customer(update, context)
        return

    if text == "🗑 Delete Customer":
        await start_delete_customer(update, context)
        return


    state = context.user_data.get("state")

    if state in [
        "search_products",
        "search_sales",
        "search_expenses"
    ]:
        await search_steps(update, context)
        return

    # Active product operations
    if state in [
        "product_name",
        "buy_price",
        "sell_price",
        "stock"
    ]:
        await add_product_steps(update, context)
        return

    if state in [
        "expense_description",
        "expense_amount"
    ]:
        await expense_steps(update, context)
        return

    if state in [
        "sale_product_id",
        "sale_quantity"
    ]:
        await new_sale_steps(update, context)
        return

    if state == "delete_product":
        await delete_product(update, context)
        return

    if state in ["edit_id", "edit_name"]:
        await edit_product(update, context)
        return

    # Main menu
    if text == "📦 Products":
        await show_products(update, context)

    elif text == "➕ Add Product":
        await start_add_product(update, context)

    elif text == "📋 View Products":
        await view_products(update, context)

    elif text == "🗑 Delete Product":
        await start_delete_product(update, context)

    elif text == "✏️ Edit Product":
        await start_edit_product(update, context)

    elif text == "🔙 Back":

        context.user_data.clear()

        await update.message.reply_text(
            "🏪 *Main Menu*",
            reply_markup=main_menu(),
            parse_mode="Markdown"
        )

    elif text == "🛒 New Sale":
        await start_new_sale(update, context)


    elif text == "💸 Expenses":
        keyboard = [
            ["➕ Add Expense", "📋 View Expenses"],
            ["🔙 Back"]
        ]

        await update.message.reply_text(
            "💸 *Expenses*\n\nChoose an option:",
            reply_markup=ReplyKeyboardMarkup(
                keyboard,
                resize_keyboard=True
            ),
            parse_mode="Markdown"
        )

    elif text == "➕ Add Expense":
        await start_expense(update, context)

    elif text == "📋 View Expenses":
        await view_expenses(update, context)

    elif text == "📈 Reports":
        await show_reports(update, context)

    elif text == "🔄 Refresh":
        await show_reports(update, context)

    elif text == "📋 Sales":
        await report_sales(update, context)

    elif text == "📦 Stock":
        await report_stock(update, context)

    elif text == "🔍 Search":
        await start_search(update, context)

    elif text == "🔎 Search Products":
        await search_products_start(update, context)

    elif text == "🔎 Search Sales":
        await search_sales_start(update, context)

    elif text == "🔎 Search Expenses":
        await search_expenses_start(update, context)

    elif text == "📁 Excel":

        await update.message.reply_text(
            "📁 Your Excel file is:\n\n"
            "`business.xlsx`",
            parse_mode="Markdown"
        )

    elif text == "⚙️ Settings":

        await update.message.reply_text(
            "⚙️ Settings coming next."
        )


# =========================
# ERROR
# =========================

async def error_handler(update, context):

    print("Error:", context.error)


# =========================
# RUN BOT
# =========================

def main():

    create_excel()

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(customer_callback, pattern=r"^customer_view:"))
    app.add_handler(CallbackQueryHandler(customer_list_callback, pattern=r"^customer_list$"))
    app.add_handler(CallbackQueryHandler(sale_callback, pattern=r"^(sale_|sale_)"))

    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            buttons
        )
    )

    app.add_error_handler(error_handler)

    print("🤖 Business Manager Bot is running...")

    app.run_polling()


if __name__ == "__main__":
    main()
