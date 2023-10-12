import xlsxwriter
import sys
import random

if len(sys.argv) != 1:
  sys.exit("Usage: python3 generator_excel.py")

options = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]

# user input
num_serial = int(input("總共生成幾組流水號+序號? "))
while num_serial <= 0:
  num_serial = int(input("總共生成幾組流水號+序號(需至少生成一組)? "))
serial_digits = int(input("一組序號中，共有幾位亂碼? "))
while serial_digits < 2:
  serial_digits = int(input("一組序號中，共有幾位亂碼(需至少有兩位亂碼)? "))
prefix = input("流水號開頭? ")
title_digits = int(input("共有幾位流水號數字? "))
while title_digits < len(str(num_serial)):
  title_digits = int(input("共有幾位流水號數字(數字不能低於總序號位數)? "))
exclude_char = input("需要排除哪些數字、字母? ").strip()
len_exclude_char = len(exclude_char)
while len_exclude_char >= 36:
  exclude_char = input("需要排除哪些數字、字母(需至少保留一個數字/字母)? ").strip()
  len_exclude_char = len(exclude_char)
remain_option_chars = len(options) - len_exclude_char
xlsx_file_name = input("Excel命名: ")


for char in exclude_char.strip():
  # as if the character is in alphabets
  if char.capitalize() in options:
    options.remove(char.capitalize())
  # as if the character is in numbers
  elif char in options:
    options.remove(char)
serial_nums = []

# generate n pairs of random serial numbers
for i in range(num_serial):
  serial_str = ''
  # loop m times to create a new random numbers
  for _ in range(serial_digits):
    rand_char = options[random.randint(0, 35 - len_exclude_char)]
    serial_str += rand_char
  serial_nums.append(serial_str)

# create hashmaps
hashmaps = [None] * remain_option_chars * remain_option_chars

# loop all serial numbers to create hashmaps
for j in range(num_serial):
  total = 0
  for k in range(serial_digits):
    # calculate hash
    total += ord(serial_nums[j][k])
  # add to hashmaps
  hash = total % len(hashmaps)
  if hashmaps[hash] != None:
    hashmaps[hash].append(serial_nums[j])
  else:
    hashmaps[hash] = [serial_nums[j]]

# check for any possible duplicate
isDuplicated = False
for element in hashmaps:
  # detect for duplicate
  if element != None:
    if len(element) > 1:
      dup_free = set(element)
      if len(dup_free) != len(element):
        isDuplicated = True
        # allocate data using hashmaps, sort by first two char index in options to find duplicate within same ASCII value for all digits
        dup_hashmaps = [None] * 3536
        for item in element:
          dup_hash_str = str(options.index(item[0])) + str(options.index(item[1]))
          dup_hash = int(dup_hash_str)
          if dup_hashmaps[dup_hash] != None:
            dup_hashmaps[dup_hash].append(item)
          else:
            dup_hashmaps[dup_hash] = [item]
        
        for dup_item in dup_hashmaps:
          if dup_item != None:
            if len(set(dup_item)) != len(dup_item):
              print(dup_item)
              print(f"偵察到重複序號: {isDuplicated}")

# create an empty xlsx file
workbook = xlsxwriter.Workbook(xlsx_file_name + ".xlsx")
worksheet = workbook.add_worksheet()
row = 0

# start writing serial numbers into the empty xlsx sheet
for i in range(num_serial):
  row_title = prefix + str(i + 1).zfill(title_digits)
  worksheet.write(row, 0, row_title)
  worksheet.write(row, 1, serial_nums[i])
  row += 1

workbook.close()
