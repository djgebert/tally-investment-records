# from datetime import datetime

# def sort_my_list(last_date, a_list):
#   last_date = datetime.strptime("01/10/2020", "%d/%m/%Y")
#   a_list.sort()
#   return last_date
  
# a_dict = {"fred":[3,7,5], "jim":[9,7,0]}

# my_date = datetime.strptime("30/09/2020","%d/%m/%Y")
# my_date = sort_my_list(my_date, a_dict["jim"])

# print(my_date, a_dict)

def add_an_item(a_list):
  a_list.append(7)

my_list = [4,7,0]
add_an_item(my_list)

print(my_list)

