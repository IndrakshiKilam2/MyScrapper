import json
import pymsgbox
import openpyxl
from playwright.sync_api import sync_playwright


def scrape_books_homepage():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://books.toscrape.com")

        
        page.wait_for_selector("article.product_pod h3 a")

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


def display_books_message_box():
    result = scrape_books_homepage()
    msg = f"Page Title:\n{result['page_title']}\n\nTop 10 Books:\n" + \
          "\n".join(result['book_titles'][:10])
    pymsgbox.alert(msg, "Books To Scrape - Homepage Data")


def save_books_to_excel():
    result = scrape_books_homepage()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Books Data"
    ws.append(["Book Titles"])
    for title in result["book_titles"]:
        ws.append([title])
    wb.save("books_data.xlsx")
    pymsgbox.alert("Data saved to 'books_data.xlsx'", "Save Successful")


def scrape_with_config():
    try:
        with open("config.json", "r") as f:
            config = json.load(f)
        result = scrape_orangehrm(config["username"], config["password"], headless=False)
        msg = f"Dashboard Title:\n{result['dashboard_title']}\n\nUser Menu:\n{result['user_menu']}"
        pymsgbox.alert(msg, "Scraped Data (From Config)")
    except FileNotFoundError:
        pymsgbox.alert("config.json not found.", "Error")
    except KeyError:
        pymsgbox.alert("config.json must contain 'username' and 'password'.", "Error")


def scrape_with_user_input():
    username = input("Enter username: ")
    password = input("Enter password: ")  
    result = scrape_orangehrm(username, password, headless=False)
    msg = f"Dashboard Title:\n{result['dashboard_title']}\n\nUser Menu:\n{result['user_menu']}"
    pymsgbox.alert(msg, "Scraped Data (From User Input)")


def main():
    while True:
        print("\nPlease select the option:")
        print("1. Screen scrap data display in message box (Books To Scrape)")
        print("2. Screen scrap data save it in Excel file (Books To Scrape)")
        print("3. Screen scrap data with user id and password from config (OrangeHRM)")
        print("4. Screen scrap data with user input (OrangeHRM)")
        print("5. Exit")

        choice = input("Enter your choice (1-5): ").strip()

        if choice == '1':
            display_books_message_box()
        elif choice == '2':
            save_books_to_excel()
        elif choice == '3':
            scrape_with_config()
        elif choice == '4':
            scrape_with_user_input()
        elif choice == '5':
            print("👋 Exiting. Goodbye!")
            break
        else:
            print("❌ Invalid input. Please enter a number between 1 and 5.")


if __name__ == "__main__":
    main()
