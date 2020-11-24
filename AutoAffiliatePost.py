import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import json
import requests
import csv
import os
import random
import string
from urllib.parse import urljoin
from datetime import datetime
import urllib3
from urllib3.exceptions import InsecureRequestWarning
from selenium import webdriver
import chromedriver_binary
import time
from selenium.webdriver.common.keys import Keys
import sys


class App():
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AutoAffiliatePost")
        self.root.geometry("1080x710")
        # ブログ
        blog_label = tk.Label(text="ブログ", font=("bold", 18))
        blog_label.place(x=10, y=20)
        style = ttk.Style(self.root)
        style.configure("my.Treeview", rowheight=40,
                        bordercolor="#ffc61e",)
        tree = ttk.Treeview(self.root, height="5", style="my.Treeview")
        tree.place(x=10, y=60)
        # 列インデックスの作成
        tree["columns"] = (1, 2, 3)
        # 表スタイルの設定(headingsはツリー形式ではない、通常の表形式)
        tree["show"] = "headings"
        # 各列の設定(インデックス,オプション(今回は幅を指定))
        tree.column(1, width=240)
        tree.column(2, width=240)
        tree.column(3, width=240)
        # 各列のヘッダー設定(インデックス,テキスト)
        tree.heading(1, text="ドメイン")
        tree.heading(2, text="ユーザー名")
        tree.heading(3, text="パスワード")
        # 値の追加・編集
        log_label = tk.Label(text="値の追加・編集", font=("bold", 18))
        log_label.place(x=750, y=20)
        input_log = tk.Entry(state='disabled')
        input_log.place(x=750, y=60, width="320", height="250")
        # 値の追加・編集 ドメイン
        input_domain = tk.Entry()
        input_domain.place(x=760, y=92, width=300)
        input_domain_label = tk.Label(text="ドメイン")
        input_domain_label.place(x=760, y=70)
        # 値の追加・編集 ユーザー名
        input_username = tk.Entry()
        input_username.place(x=760, y=142, width=300)
        input_username_label = tk.Label(text="ユーザー名")
        input_username_label.place(x=760, y=120)
        # 値の追加・編集 パスワード
        input_password = tk.Entry()
        input_password.place(x=760, y=192, width=300)
        input_password_label = tk.Label(text="パスワード")
        input_password_label.place(x=760, y=170)
        # コンボボックス
        blog_list = []

        def combobox_selected(self, event):
            print(c.get())
        variable = tk.StringVar(self.root)
        c = ttk.Combobox(self.root, values=blog_list,
                         textvariable=variable)
        c.place(x=90, y=370, width=490)
        c.bind("<<ComboboxSelected>>", combobox_selected)
        # コンボボックスに値をセット

        def set_value_to_c(clist):
            c = ttk.Combobox(self.root, values=clist,
                             textvariable=variable)
            c.place(x=90, y=370, width=490)
            num = len(clist) - 1
            return
        # csvがあったら事前に読み込み
        if os.path.isfile("../blog_data.csv"):
            with open("../blog_data.csv") as f:
                reader = csv.reader(f)
                for row in reader:
                    tree.insert("", "end", values=row)
                    blog_list.append(row[0])
                    set_value_to_c(blog_list)
        # 値の追加

        def add_data():
            # validation
            if input_domain.get() == "":
                return messagebox.showerror("エラー", "ドメインを入力してください")
            if input_username.get() == "":
                return messagebox.showerror("エラー", "ユーザー名を入力してください")
            if input_domain.get() == "":
                return messagebox.showerror("エラー", "パスワードを入力してください")
            # 表に値を挿入
            tree.insert("", "end", values=(input_domain.get(),
                                           input_username.get(), input_password.get()))
            blog_list.append(input_domain.get())
            set_value_to_c(blog_list)
            input_domain.delete(0, tk.END)
            input_username.delete(0, tk.END)
            input_password.delete(0, tk.END)
            return
        button_add = tk.Button(text="値の追加", bg='#cafddc',
                               width=29, command=add_data)
        button_add.place(x=800, y=230)
        # 値のセット

        def get_value(self):
            selcted_item = tree.focus()
            values = tree.item(selcted_item, "values")
            domain = values[0]
            username = values[1]
            password = values[2]
            input_domain.delete(0, tk.END)
            input_username.delete(0, tk.END)
            input_password.delete(0, tk.END)
            input_domain.insert(tk.END, domain)
            input_username.insert(tk.END, username)
            input_password.insert(tk.END, password)
        tree.bind("<<TreeviewSelect>>", get_value)
        # 編集の反映

        def edit_data():
            # validation
            if tree.focus() == "":
                return messagebox.showerror("エラー", "編集する行を入力してください")
            if input_domain.get() == "":
                return messagebox.showerror("エラー", "ドメインを入力してください")
            if input_username.get() == "":
                return messagebox.showerror("エラー", "ユーザー名を入力してください")
            if input_password.get() == "":
                return messagebox.showerror("エラー", "パスワードを入力してください")
            # 編集内容を表に反映
            selcted_item = tree.focus()
            data = [input_domain.get(), input_username.get(),
                    input_password.get()]
            values = tree.item(selcted_item, values=data)
        button_edit = tk.Button(
            text="編集を反映", bg='#cafddc', width=29, command=edit_data)
        button_edit.place(x=800, y=265)

        def save_data():
            with open("../blog_data.csv", "w", newline='') as myfile:
                csvwriter = csv.writer(myfile, delimiter=',')
                for row_id in tree.get_children():
                    row = tree.item(row_id)['values']
                    csvwriter.writerow(row)
                messagebox.showinfo("確認", "保存に成功しました")
        button_save = tk.Button(
            text="変更の保存", bg='#cafddc', width=15, command=save_data)
        button_save.place(x=100, y=20)
        # ポスト
        post_label = tk.Label(text="ポスト", font=("bold", 18))
        post_label.place(x=10, y=320)
        # ブログ
        for row_id in tree.get_children():
            blog_data = tree.item(row_id)['values']
            blog_list.append(blog_data[0])
        input_blog_label = tk.Label(text="ブログ：")
        input_blog_label.place(x=10, y=370)
        # キーワード
        input_keyword = tk.Entry(width=60)
        input_keyword.place(x=90, y=430)
        input_keyword_label = tk.Label(text="キーワード：")
        input_keyword_label.place(x=10, y=430)
        # Rakuten ユーザID
        input_rakuten_id = tk.Entry(show='*', width=53)
        input_rakuten_id.place(x=145, y=490)
        input_rakuten_id_label = tk.Label(text="Rakuten　ユーザID：")
        input_rakuten_id_label.place(x=10, y=490)
        # Rakuten パスワード
        input_rakuten_pw = tk.Entry(show='*', width=53)
        input_rakuten_pw.place(x=145, y=530)
        input_rakuten_pw_label = tk.Label(text="パスワード：")
        input_rakuten_pw_label.place(x=60, y=530)
        # Amazon
        input_amazon = tk.Entry(show='*', width=60)
        input_amazon.place(x=90, y=590)
        input_amazon_label = tk.Label(text="Amazon：")
        input_amazon_label.place(x=10, y=590)
        # Rakten Affiliate
        input_rakuten = tk.Entry(show='*', width=60)
        input_rakuten.place(x=90, y=630)
        input_rakuten_label = tk.Label(text="Rakuten：")
        input_rakuten_label.place(x=10, y=630)
        # Yahoo SID
        input_yahoo_SID = tk.Entry(show='*', width=25)
        input_yahoo_SID.place(x=90, y=670)
        input_yahoo_SID_label = tk.Label(text="Yahoo SID：")
        input_yahoo_SID_label.place(x=10, y=670)
        # Yahoo PID
        input_yahoo_PID = tk.Entry(show='*', width=25)
        input_yahoo_PID.place(x=370, y=670)
        input_yahoo_PID_label = tk.Label(text="PID：")
        input_yahoo_PID_label.place(x=330, y=670)
        # URLs
        input_URLs_label = tk.Label(text="URLs：")
        input_URLs_label.place(x=620, y=350)
        input_URL = tk.Text(self.root, font=("", 13), height=20)
        input_URL.place(x=680, y=350, width=390)

        def post_article(status, slug, title, content, category_ids, tag_ids, media_id):
            """
            記事を投稿して成功した場合はTrue、失敗した場合はFalseを返します。
            :param status: 記事の状態（公開:publish, 下書き:draft）
            :param slug: 記事識別子。URLの一部になる（ex. slug=aaa-bbb/ccc -> https://wordpress-example.com/aaa-bbb/ccc）
            :param title: 記事のタイトル
            :param content: 記事の本文
            :param category_ids: 記事に付与するカテゴリIDのリスト
            :param tag_ids: 記事に付与するタグIDのリスト
            :param media_id: 見出し画像のID
            :return: レスポンス
            """
            # credential and attributes
            user_ = WP_USERNAME
            pass_ = WP_PASSWORD
            # build request body
            payload = {
                "status": status,
                "slug": slug,
                "title": title,
                "content": content,
                "date": datetime.now().isoformat(),
                "categories": category_ids,
                "tags": tag_ids
            }
            if media_id is not None:
                payload['featured_media'] = media_id
            # 警告を無効化
            urllib3.disable_warnings(InsecureRequestWarning)
            # send POST request
            res = requests.post(
                urljoin(WP_URL, "wp-json/wp/v2/posts"),
                data=json.dumps(payload),
                headers={'Content-type': "application/json"},
                auth=(user_, pass_)
            )
            print('----------\n件名:「{}」res.status: {}'.format(title,
                                                             repr(res.status_code)))
            if res.status_code != 201:
                return messagebox.showerror("エラー", "投稿に失敗しました。入力する値を再度確認してみてください。")
            return res

        def randomname(n):
            randlst = [random.choice(
                string.ascii_letters + string.digits) for i in range(n)]
            return ''.join(randlst)
        # 実行

        def do_data():
            # validation
            if input_keyword.get() == "":
                return messagebox.showerror("エラー", "「キーワード」の値を入力してください")
            if input_rakuten_id.get() == "":
                return messagebox.showerror("エラー", "「Rakuten ID」の値を入力してください")
            if input_rakuten_pw.get() == "":
                return messagebox.showerror("エラー", "「Rakuten パスワード」の値を入力してください")
            if input_rakuten.get() == "":
                return messagebox.showerror("エラー", "「Rakuten」の値を入力してください")
            if input_amazon.get() == "":
                return messagebox.showerror("エラー", "「Amazon」の値を入力してください")
            if input_yahoo_SID.get() == "":
                return messagebox.showerror("エラー", "「Yahoo SID」の値を入力してください")
            if input_yahoo_PID.get() == "":
                return messagebox.showerror("エラー", "「Yahoo PID」の値を入力してください")
            for row in tree.get_children():
                blog_data = tree.item(row)['values']
                if blog_data[0] == c.get():
                    # WordPressのログイン情報をglobal宣言
                    global WP_URL, WP_USERNAME, WP_PASSWORD
                    WP_URL = blog_data[0]
                    WP_USERNAME = blog_data[1]
                    WP_PASSWORD = blog_data[2]
            with open("../post_data.csv", "w", newline='') as myfile:
                csvwriter = csv.writer(myfile, delimiter=',')
                row = [
                    input_rakuten_id.get(),
                    input_rakuten_pw.get(),
                    input_rakuten.get(),
                    input_amazon.get(),
                    input_yahoo_SID.get(),
                    input_yahoo_PID.get()
                ]
                csvwriter.writerow(row)
            slug = randomname(15)
            keyword = input_keyword.get()

            def resource_path(relative_path):
                try:
                    base_path = sys._MEIPASS
                except Exception:
                    base_path = os.path.dirname(__file__)
                return os.path.join(base_path, relative_path)

            def get_affilink(URL):
                driver = webdriver.Chrome("chromedriver.exe")
                # アフィリエイトリンク生成トップページ
                driver.get("https://affiliate.rakuten.co.jp/")
                time.sleep(12)
                driver.find_element_by_name("u").send_keys(URL)
                driver.find_element_by_xpath(
                    '//button[text()="作成"]').click()
                # ログインページ
                driver.find_element_by_name(
                    "u").send_keys(input_rakuten_id.get())
                time.sleep(10)
                driver.find_element_by_name(
                    "p").send_keys(input_rakuten_pw.get())
                driver.find_element_by_xpath(
                    '//*[@id="loginInner"]/p[1]/input').click()
                # アフィリエイトリンク生成ページ
                time.sleep(10)
                pulldowns = driver.find_elements_by_xpath(
                    '//span[@id="look_here"]')
                pulldowns[1].click()
                time.sleep(5)
                down_element = driver.find_element_by_xpath(
                    '//*[@id="caption_pos_container"]/label[3]')
                down_element.click()
                affilink1 = driver.find_element_by_id(
                    "codebox").get_attribute("value")
                driver.get(URL)
                # 商品ページ
                time.sleep(14)
                try:
                    catch_copy = driver.find_element_by_class_name(
                        "catch_copy").text
                except:
                    catch_copy = ""
                    pass
                try:
                    item_name = driver.find_element_by_class_name(
                        "item_name").text
                except:
                    item_name = ""
                    pass
                try:
                    item_price = driver.find_element_by_class_name(
                        "price2").text
                except:
                    item_price = ""
                    pass
                try:
                    shipping_price = driver.find_element_by_class_name(
                        "dsf-shipping-cost").text
                except:
                    shipping_price = ""
                    pass
                price = item_price + "+" + shipping_price
                driver.find_element_by_xpath(
                    '//button[contains(text(), "商品レビューを見る")]').click()
                # レビューページへ
                time.sleep(13)
                try:
                    review_ave = driver.find_element_by_class_name(
                        "revEvaNumber").text
                except:
                    review_ave = ""
                try:
                    review_num = driver.find_element_by_class_name(
                        "Count").text
                except:
                    review_num = ""
                review = "☆評価" + review_ave + "（レビュー数：" + review_num + "件）"
                data = [catch_copy, affilink1, item_name, price, review]
                driver.close()
                return data

            def get_commonlink():
                driver = webdriver.Chrome()
                # 商品検索ページ
                driver.get('https://www.rakuten.co.jp/')
                time.sleep(12)
                driver.find_element_by_name("sitem").send_keys(keyword)
                driver.find_element_by_name("sitem").send_keys(Keys.ENTER)
                # 商品検索結果表示ページ
                time.sleep(15)
                img = driver.find_element_by_class_name(
                    "_verticallyaligned").get_attribute("src")
                desc = driver.find_element_by_class_name(
                    "_verticallyaligned").get_attribute("alt")
                link = driver.find_element_by_xpath(
                    '//a[@title="' + desc + '"]').get_attribute("href")
                driver.get(link)
                # 商品ページ
                time.sleep(17)
                item_URL = driver.current_url
                common_link = ' <div style="display: flex; display: -ms-flexbox; display: -webkit-box; display: -webkit-flex; width: 100%; height: 250; border: solid 1px #a1a499; background-color: #FFFFFF; box-sizing: border-box; box-shadow: 0px 2px 5px rgba(0,0,0,0.1); padding: 15px;">\
									<div style="padding: 0; display: flex; vertical-align: middle; justify-content: center;valign-items: center;">\
										<a rel="nofollow" href="' + item_URL + '">\
											<img src="' + img + '" ' + 'width="200" height="200" style="border: none;">\
										</a>\
									</div>\
									<div style="width: calc(100% - 175px);">\
										<div style= "padding:5px;">\
											<a rel="nofollow" href="' + item_URL + '" style="color: #333; text-decoration: none; padding:5px;">' + desc + '</a>\
										</div>\
										<ul style="border: none; list-style-type: none; display: inline-flex; display: -ms-inline-flexbox; display: -webkit-inline-flex; -ms-flex-wrap: wrap; flex-wrap: wrap; margin: 10px auto; padding: 0; width: 100%;">\
											<li class="amazonlink" style="background: #f6a306; list-style: none; padding: 0 18px;">\
												<a rel="nofollow" href="https://www.amazon.co.jp/s?k=' + keyword + '&creative=6339&linkCode=ure&tag=' + input_amazon.get() + '" ' + 'style="font-weight: bold; color: #FFFFFF; text-decoration: none; cursor: pointer; cursor: hand;">Amazon</a>\
											</li>\
											<li class="rakutenlink" style="background: #cf4944; list-style: none; padding: 0 18px;">\
												<a rel="nofollow" href="' + item_URL + '" ' + 'style="font-weight: bold; color: #FFFFFF; text-decoration: none; cursor: pointer; cursor: hand;">楽天市場</a>\
											</li>\
											<li class="yahoolink" style="background: #51a7e8; list-style: none; font-weight: bold; color: #FFFFFF; padding: 0 18px;">\
												<a rel="nofollow" href="https://shopping.yahoo.co.jp/search?p=' + keyword + '&sc_e=afvc_shp_3538913" style="font-weight: bold; color: #FFFFFF; text-decoration: none; cursor: pointer; cursor: hand;">Yahooショッピング</a>\
											</li>\
										</ul>\
									</div>\
								</div>'
                driver.close()
                return common_link
            URLs = (input_URL.get("1.0", "end")).splitlines()
            contents = ""
            first_loop = True
            for link in URLs:
                if first_loop:
                    common_link = get_commonlink()
                    first_loop = False
                if link == "":
                    break
                data = get_affilink(link)
                catch_copy = data[0]
                affilink = data[1]
                item_name = data[2]
                price = data[3]
                review = data[4]
                content = "<h2></h2><br>" + "<div style='color: #e60033; font-weight: bold;'>" + catch_copy + "</div>" + affilink + item_name + \
                    "<br><span style='background: linear-gradient(transparent 80%, yellow 80%);'>" + price + "<br>" + review + "<br></span>" + \
                    "<p><span style='font-weight: bold;'>価格は常に変動します。</span><br><span style='font-weight: bold;'>最安値比較はコチラから！</span><br>↓　↓　↓<br></p>" + common_link
                contents = contents + content
            # 記事を下書き投稿する（'draft'ではなく、'publish'にすれば公開投稿できます。）
            post_article('draft', slug, keyword, contents,
                         category_ids=[], tag_ids=[], media_id=None)
        # 実行ボタン
        button_do = tk.Button(text="▶ 実行", bg='#cafddc',
                              width=15, command=do_data)
        button_do.place(x=100, y=320)
        # ポスト クリアボタン

        def clear_postdata():
            input_keyword.delete(0, tk.END)
            input_rakuten_id.delete(0, tk.END)
            input_rakuten_pw.delete(0, tk.END)
            input_amazon.delete(0, tk.END)
            input_rakuten.delete(0, tk.END)
            input_yahoo_PID.delete(0, tk.END)
            input_yahoo_SID.delete(0, tk.END)
            return
        post_clear = tk.Button(text="値のクリア", bg='#F99191',
                               width=15, command=clear_postdata)
        post_clear.place(x=240, y=320)
        # ブログ クリアボタン

        def value_reset(b_list):
            b_list.clear()
            return

        def clear_blogdata():
            selected_item = tree.selection()[0]
            tree.delete(selected_item)
            value_reset(blog_list)
            for row_id in tree.get_children():
                blog_data = tree.item(row_id)['values']
                blog_list.append(blog_data[0])
            set_value_to_c(blog_list)
            input_domain.delete(0, tk.END)
            input_username.delete(0, tk.END)
            input_password.delete(0, tk.END)
            return
        blog_clear = tk.Button(text="行の削除", bg='#F99191',
                               width=15, command=clear_blogdata)
        blog_clear.place(x=240, y=20)
        if os.path.isfile("../post_data.csv"):
            with open("../post_data.csv") as f:
                reader = csv.reader(f)
                for row in reader:
                    input_rakuten_id.insert(tk.END, row[0])
                    input_rakuten_pw.insert(tk.END, row[1])
                    input_amazon.insert(tk.END, row[2])
                    input_rakuten.insert(tk.END, row[3])
                    input_yahoo_SID.insert(tk.END, row[4])
                    input_yahoo_PID.insert(tk.END, row[5])
        self.root.mainloop()

    def quit(self):
        self.root.destroy()
        return


app = App()
