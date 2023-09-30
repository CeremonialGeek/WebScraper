import tkinter as tk
from tkinter import filedialog, messagebox
import requests
from bs4 import BeautifulSoup
import csv
import os
import re

def browse_upload_path():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        part_text.delete("1.0", "end")
        open_path_entry.delete(0, "end")
        open_path_entry.insert("end", file_path)

def write_to_file(save_path, results):
    try:
        with open(save_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Part Number', 'Part Status PlaymoDB', 'PlaymoDB Part Info', 'Colour', 'PlaymoDB Sets',
                             'Part Status Playmobil.com', 'Playmobil.com Name', 'Price', 'URL'])
            for result_arr in results:
                result_arr[-1] = f'=HYPERLINK("{result_arr[-1]}","click here for more")'
                writer.writerow(result_arr)

        messagebox.showinfo("Success", f"Saved to {save_path}")
        part_text.delete("1.0", "end")
    except Exception as e:
        messagebox.showinfo("File Not Saved", f"Error!: {e}")

def browse_save_path():
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    save_path_entry.delete(0, "end")
    save_path_entry.insert("end", save_path)

def scrape():
    part_input = part_text.get("1.0", "end").strip()
    upload_path = open_path_entry.get()
    if upload_path.endswith('.csv'):
        with open(upload_path, newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile)
            part_arr = [number.replace(" ", "") for row in reader for number in row]
            print(part_arr)
    else:
        part_arr = re.split(r"[\s,\n]+", part_input)

    if len(part_arr) == 0 or (len(part_arr) == 1 and not part_arr[0]):
        messagebox.showerror("Error", "Enter At Least 1 Part Number!")
        return

    results = []

    for p in part_arr:
        result_arr = []
        # Getting Website Elements
        playmobil_url = "https://www.playmobil.co.uk/on/demandware.store/Sites-GB-Site/en_GB/CustomSearch-Spareparts?q=" + p
        playmobil_req = requests.get(playmobil_url)
        playmobil_soup = BeautifulSoup(playmobil_req.text, 'html.parser')
        playmobil_part_div = playmobil_soup.find('div', class_='col-s-6 col-xs-12 sparepart-content')

        playmodb_url = "https://playmodb.org/cgi-bin/showpart.pl?partnum=" + p
        playmodb_req = requests.get(playmodb_url)
        playmodb_soup = BeautifulSoup(playmodb_req.text, 'html.parser')
        part_h2 = playmodb_soup.find('h2')
        part_span = playmodb_soup.find('span')
        sets_playmodb = playmodb_soup.find_all('p')
        div_elements = playmodb_soup.find_all('div', class_='setlineup')
        result_arr.append(p)

        if part_h2 is not None:
            part_info = part_h2.text.split("PlaymoDB Part Info for")[-1].strip()
            if part_info.lower() == "(unknown)":
                result_arr.append('Not Included')
            else:
                result_arr.append('Included')
            result_arr.append(part_info)
        else:
            result_arr.append('Not Included')
            result_arr.append('Unknown')

        if part_span is not None:
            part_span_txt = part_span.get_text(strip=True)
            if "Colour:" in part_span_txt:
                part_span_split = part_span_txt.split("Colour:")[-1].strip()
                result_arr.append(part_span_split)
            else:
                result_arr.append('Colour Not Found')
        else:
            result_arr.append('Colour Not Found')

        # Logic For Getting Set Numbers
        setnumbers_arr = []
        for element in sets_playmodb:
            if "Sets containing this part (confirmed):" in element.text:
                if "Permalink in HTML:" not in element.text:
                    element_not_split = element.get_text(strip=True)
                    element_split = element_not_split.split("Sets containing this part (confirmed):")[-1]
                    print(element_split)
                    setnumbers_arr.append(element_split)
                else:
                    for div_element in div_elements:
                        numbertext = div_element.get_text(strip=True)
                        setnumber = re.search(r'\d+', numbertext)
                        if setnumber:
                            setnumbers_arr.append(setnumber.group())

        if not setnumbers_arr:
            setnumbers_arr.append("Not Included!")
        result_arr.append(', '.join(setnumbers_arr))

        if playmobil_part_div is not None:
            part_name = playmobil_part_div.find('h2').get_text(strip=True)
            part_price = playmobil_part_div.find('span').get_text(strip=True)
            result_arr.append('Available from Playmobil.com')
            result_arr.append(part_name)
            result_arr.append(part_price)
            result_arr.append(playmobil_url)

        else:
            result_arr.append('Unavailable from Playmobil.com')
            result_arr.append('Not Found')
            result_arr.append('Not Found')

        results.append(result_arr)

    save_path = save_path_entry.get()
    if save_path:
        print(results)
        write_to_file(save_path, results)


# Gui Elements
window = tk.Tk()
window.title("Playmobil Parts Scraper")
window.geometry("500x500")
window.resizable(False, False)
window.configure(background='white')

title_label = tk.Label(window, text="PLAYMOBIL PART SCRAPER", font=("Corbel", 14, "bold"), bg="white")
title_label.pack(pady=10)

instruction_label = tk.Label(window, text="Enter or Paste Part Numbers:", font=("Corbel", 12), bg="white")
instruction_label.pack(pady=10)

part_text = tk.Text(window, font=("Calibri", 12), height=5)
part_text.pack(pady=10)

save_path_frame = tk.Frame(window, bg="white")
save_path_frame.pack(pady=10)

save_path_label = tk.Label(save_path_frame, text="Save Location:", font=("Corbel", 12), bg="white")
save_path_label.pack(side="left")

save_path_entry = tk.Entry(save_path_frame, font=("Corbel", 12))
save_path_entry.pack(side="left", padx=5)

browse_button = tk.Button(save_path_frame, text="Browse", font=("Corbel", 12), width=10, command=browse_save_path)
browse_button.pack(side="left")

open_path_frame = tk.Frame(window, bg="white")
open_path_frame.pack(pady=10)

open_path_label = tk.Label(open_path_frame, text="Upload Part Numbers:", font=("Corbel", 12), bg="white")
open_path_label.pack(side="left")

open_path_entry = tk.Entry(open_path_frame, font=("Corbel", 12))
open_path_entry.pack(side="left", padx=5)

upload_button = tk.Button(open_path_frame, text="Upload", font=("Corbel", 12), width=10, command=browse_upload_path)
upload_button.pack(side="left")

scrape_button = tk.Button(window, text="Scrape", font=("Corbel", 12), width=10, command=scrape)
scrape_button.pack(pady=10)

made_by_label = tk.Label(window, text="Made by Hannah B", font=("Fixedsys", 10), bg="white")
made_by_label.pack(pady=5)

version_label = tk.Label(window, text="Version 1.3", font=("Terminal", 10), bg="white")
version_label.pack(pady=5)
playmobilcopyright_label = tk.Label(window, text="Playmobil Copyright of: Brandstätter Group", bg="white",
                                    fg="blue", font=("Terminal", 10))
playmobilcopyright_label.pack(pady=5)

window.mainloop()
