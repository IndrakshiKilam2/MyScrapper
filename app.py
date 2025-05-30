import json
import pymsgbox
import openpyxl
from flask import Flask, request, render_template_string, send_file
from playwright.sync_api import sync_playwright
from io import BytesIO

app = Flask(__name__)


def scrape_books_homepage():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://books.toscrape.com")

        page_title = page.title()
        books = page.query_selector_all("article.product_pod h3 a")
        book_titles = [book.get_attribute("title") for book in books]
        browser.close()

        return {
            "page_title": page_title,
            "book_titles": book_titles
        }


def scrape_orangehrm(username, password, headless=False):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()
        page.goto("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")

        page.wait_for_selector('input[name="username"]')
        page.fill('input[name="username"]', username)
        page.fill('input[name="password"]', password)
        page.click('button[type="submit"]')

        page.wait_for_selector('h6.oxd-text.oxd-text--h6.oxd-topbar-header-breadcrumb-module')
        dashboard_title = page.inner_text('h6.oxd-text.oxd-text--h6.oxd-topbar-header-breadcrumb-module')

        page.click('p.oxd-userdropdown-name')
        user_menu = page.inner_text('ul.oxd-dropdown-menu')

        browser.close()
        return {
            "dashboard_title": dashboard_title,
            "user_menu": user_menu
        }

# Home Page
UI_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<title>MYSCRAPER</title>

<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Screen Scraping Menu</title>
<style>
    h1 {
        text-align: center;
        color: #34495e;
        font-size: 3rem;
        margin-bottom: 30px;
        letter-spacing: 3px;
        font-weight: bold;
    }
    body {
        background: #f0f2f5;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
        margin: 0;
    }
    .container {
        background: white;
        padding: 30px 40px;
        border-radius: 10px;
        box-shadow: 0 8px 16px rgba(0,0,0,0.15);
        max-width: 400px;
        width: 100%;
        text-align: center;
    }
    h2 {
        margin-bottom: 24px;
        color: #333;
    }
    button {
        background-color: #007bff;
        border: none;
        color: white;
        padding: 12px 20px;
        margin: 10px 0;
        border-radius: 6px;
        font-size: 16px;
        cursor: pointer;
        width: 100%;
        transition: background-color 0.3s ease;
    }
    button:hover {
        background-color: #0056b3;
    }
    a {
        display: inline-block;
        margin-top: 20px;
        text-decoration: none;
        color: #007bff;
        font-weight: 600;
    }
    a:hover {
        text-decoration: underline;
    }
    form label {
        display: block;
        margin: 12px 0 6px;
        font-weight: 600;
        color: #444;
        text-align: left;
    }
    form input {
        width: 100%;
        padding: 10px 12px;
        border-radius: 6px;
        border: 1px solid #ccc;
        font-size: 15px;
    }
</style>
</head>
<body>
<div class="container">
    <h1>MYSCRAPER</h1>
    <h2>Please select an option</h2>
    <form method="post" action="/scrape">
        <button name="option" value="1" type="submit">1. Screen scrap Books to Scrape homepage & display in message box</button>
        <button name="option" value="2" type="submit">2. Screen scrap Books to Scrape homepage & save it in Excel file</button>
        <button name="option" value="3" type="submit">3. Screen scrap OrangeHRM with credentials from config</button>
        <button name="option" value="4" type="submit">4. Screen scrap OrangeHRM with user input</button>
        <button name="option" value="5" type="submit" style="background:#dc3545;">5. Exit</button>
    </form>
</div>
</body>
</html>
"""

# User Input 
USER_INPUT_FORM = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Login Credentials Input</title>
<style>
    body {
        background: #f0f2f5;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
        margin: 0;
    }
    .container {
        background: white;
        padding: 30px 40px;
        border-radius: 10px;
        box-shadow: 0 8px 16px rgba(0,0,0,0.15);
        max-width: 400px;
        width: 100%;
        text-align: center;
    }
    h1 {
        text-align: center;
        color: #34495e;
        font-size: 3rem;
        margin-bottom: 30px;
        letter-spacing: 3px;
        font-weight: bold;
    }
    h2 {
        margin-bottom: 24px;
        color: #333;
    }
    form label {
        display: block;
        margin: 12px 0 6px;
        font-weight: 600;
        color: #444;
        text-align: left;
    }
    form input {
        width: 100%;
        padding: 10px 12px;
        border-radius: 6px;
        border: 1px solid #ccc;
        font-size: 15px;
    }
    button {
        margin-top: 20px;
        background-color: #007bff;
        border: none;
        color: white;
        padding: 12px 20px;
        border-radius: 6px;
        font-size: 16px;
        cursor: pointer;
        width: 100%;
        transition: background-color 0.3s ease;
    }
    button:hover {
        background-color: #0056b3;
    }
    a {
        display: inline-block;
        margin-top: 20px;
        text-decoration: none;
        color: #007bff;
        font-weight: 600;
    }
    a:hover {
        text-decoration: underline;
    }
</style>
</head>
<body>

<div class="container">
    <h2>Enter Login Credentials</h2>
    <form method="post" action="/scrape_user_input">
        <label for="username">Username:</label>
        <input id="username" name="username" type="text" required autocomplete="username" />
        
        <label for="password">Password:</label>
        <input id="password" name="password" type="password" required autocomplete="current-password" />
        
        <button type="submit">Submit</button>
    </form>
    <a href="/">Back to menu</a>
</div>
</body>
</html>
"""

@app.route("/", methods=["GET"])
def index():
    return render_template_string(UI_HTML)

@app.route("/scrape", methods=["POST"])
def scrape():
    option = request.form.get("option")

    if option == "1":
        try:
            result = scrape_books_homepage()
            msg = f"Page Title:\n{result['page_title']}\n\nTop 10 Books:\n" + "\n".join(result['book_titles'][:10])
            pymsgbox.alert(msg, "Books To Scrape - Homepage Data")
            return render_template_string(UI_HTML)
        except Exception as e:
            return f"<p style='color:red;'>Error: {e}</p><a href='/'>Go Back</a>"

    elif option == "2":
        try:
            result = scrape_books_homepage()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Books Data"
            ws.append(["Book Titles"])
            for title in result["book_titles"]:
                ws.append([title])
            file_stream = BytesIO()
            wb.save(file_stream)
            file_stream.seek(0)
            return send_file(
                file_stream,
                download_name="books_data.xlsx",
                as_attachment=True,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            return f"<p style='color:red;'>Error: {e}</p><a href='/'>Go Back</a>"

    elif option == "3":
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
            result = scrape_orangehrm(config["username"], config["password"], headless=True)
            msg = f"Dashboard Title:\n{result['dashboard_title']}\n\nUser Menu:\n{result['user_menu']}"
            pymsgbox.alert(msg, "Scraped Data (From Config)")
            return render_template_string(UI_HTML)
        except FileNotFoundError:
            return "<p style='color:red;'>Error: config.json not found.</p><a href='/'>Go Back</a>"
        except KeyError:
            return "<p style='color:red;'>Error: config.json must contain 'username' and 'password'.</p><a href='/'>Go Back</a>"
        except Exception as e:
            return f"<p style='color:red;'>Error: {e}</p><a href='/'>Go Back</a>"

    elif option == "4":
        return render_template_string(USER_INPUT_FORM)

    elif option == "5":
        return "<p style='font-size:1.2rem;'>👋 Goodbye! Close this tab to exit.</p>"

    else:
        return "<p style='color:red;'>Invalid option.</p><a href='/'>Go Back</a>"

@app.route("/scrape_user_input", methods=["POST"])
def scrape_user_input():
    username = request.form.get("username")
    password = request.form.get("password")
    if not username or not password:
        return "<p style='color:red;'>Please provide both username and password.</p><a href='/'>Go Back</a>"

    try:
        result = scrape_orangehrm(username, password, headless=True)
        msg = f"Dashboard Title:\n{result['dashboard_title']}\n\nUser Menu:\n{result['user_menu']}"
        pymsgbox.alert(msg, "Scraped Data (From User Input)")
        return render_template_string(UI_HTML)
    except Exception as e:
        return f"<p style='color:red;'>Error: {e}</p><a href='/'>Go Back</a>"

if __name__ == "__main__":
    app.run(debug=True)
