import re
import time
import random
import psutil
from pywinauto import Application
from selenium import webdriver
from selenium.webdriver.common.by import By

print("Opening Browser")
driver = webdriver.Chrome()
print(type(driver))
#driver.maximize_window()
driver.get("https://www.wordleunlimited.com/")
time.sleep(5)
print("Connecting pywinauto")
process = psutil.Process(driver.service.process.pid)

app = Application()
app.connect(process=process.children()[0].pid, timeout=10)
form = app.top_window()
print("Connected")
time.sleep(0.5)

wins = 0
losses = 0
streak = 0

def calculate():
    tmp_active_set = active_words.copy()

    for word in tmp_active_set:
        # removes words that have incorrect letters in them
        for letter in incorrect:
            if letter in word:
                try:
                    active_words.remove(word)
                except:
                    pass
                break

        # removes words that don't have correct letters in them
        for letter1 in correct:
            if letter1 not in word:
                try:
                    active_words.remove(word)
                except:
                    pass
                break
        
        # removes words that don't have correct letter position in them
        for check in correct_pos:
            if check[0] not in word[check[1]]:
                try:
                    active_words.remove(word)
                except:
                    pass
                break

        # removes words that have incorrect letter positioning in them
        for check1 in incorrect_pos:
            if check1[0] == word[check1[1]]:
                try:
                    active_words.remove(word)
                except:
                    pass
                break

while True:
    alphabet = {'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z'}
    used = set()
    unused = alphabet
    correct = set()
    incorrect = set()
    correct_pos = set()
    incorrect_pos = set()
    active_words = set()

    # put all dictionary words into set
    word_file = open('words_5.txt', 'r')
    word_lines = word_file.readlines()
    for word in word_lines:
        active_words.add(word.strip().lower())
    word_file.close()
    words_set = active_words.copy()

    letter_pos = 0

    count = 0

    while True:
        if count == 0:
            word = "vitae"
        elif count == 1:
            word = "hymns"
        else:
            try:
                word = random.choice(tuple(active_words))
            except:
                word = "based"

        #word = str(input("Word: ")).strip().lower()
        if len(word) == 5 and word in words_set:
            count += 1
            form.type_keys(word)
            form.type_keys('{ENTER}')
            time.sleep(0.5)

            pageSource = driver.find_element(By.CLASS_NAME, "Game").get_attribute("outerHTML")

            time.sleep(0.5)

            if "Winner!" in re.findall('<p id=\"hint(.*?)</p>', pageSource)[0]:
                wins += 1
                streak += 1
                print(f"\nWins: {wins}\nLosses: {losses}\nStreak: {streak}")
                time.sleep(5)
                form.type_keys('{ENTER}')
                break
            elif "You lost!" in re.findall('<p id=\"hint(.*?)</p>', pageSource)[0]:
                losses += 1
                streak = 0
                print(f"\nWins: {wins}\nLosses: {losses}\nStreak: {streak}")
                time.sleep(5)
                form.type_keys('{ENTER}')
                break
            elif "Not a valid word" in re.findall('<p id=\"hint(.*?)</p>', pageSource)[0]:
                form.type_keys('{BACKSPACE}')
                form.type_keys('{BACKSPACE}')
                form.type_keys('{BACKSPACE}')
                form.type_keys('{BACKSPACE}')
                form.type_keys('{BACKSPACE}')
                active_words.remove(word)
                continue

            occurances = re.findall('<div class=\"Row-letter(.*?)</div>', pageSource)
            for occur in range(5):
                if "letter-elsewhere" in occurances[occur + letter_pos]:
                    incorrect_pos.add((word[occur], occur))
                    correct.add(word[occur])
                elif "letter-correct" in occurances[occur + letter_pos]:
                    correct_pos.add((word[occur], occur))
                    correct.add(word[occur])
                elif "letter-absent" in occurances[occur + letter_pos]:
                    if word.count(word[occur]) == 1:
                        incorrect.add(word[occur])

            for letter in word:
                if letter not in used:
                    used.add(letter)
                    unused.remove(letter)

            letter_pos += 5
            
            calculate()

            print(f'Used: {used}')
            print(f'Unused: {unused}')
            print(f'Correct: {correct}')
            print(f'Incorrect: {incorrect}')    
            print(f'Correct Pos: {correct_pos}')
            print(f'Incorrect Pos: {incorrect_pos}')    

            if len(active_words) < 100:
                print(active_words)
            else:
                print(len(active_words))

driver.quit()

"""
# code used for generating dictionary

words = open('words.txt', 'r')
lines = words.readlines()
words.close()

for word in lines:
    if len(word.strip()) == 5:
        with open('words_5.txt', 'a') as word_file:
            word_file.write(f'{word.strip().lower()}\n')
        word_file.close()
"""

