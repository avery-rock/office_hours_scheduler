from excel_things import open_excel
import re
file_name = 'C201_C286 Scheduling Poll (Responses).xlsx'
xl = open_excel(file_name)
titles = xl[0][:]
for i in range(len(titles)):
	titles[i] = str(titles[i]).replace('At which times can you attend office hours or help sessions?', '').strip()

print(titles)
times = titles[1:-1]
print(repr(times))

