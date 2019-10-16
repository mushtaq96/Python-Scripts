# This script helps us to google something through the CLI and the end output
# is a display of all the top 5 tabs of the required results for the search terms typed in

# ex- python TimeSaver.py Indian cricket
# result - the top 5 tabs opened in chrome

# reason?
# saves the time in using the mouse to right click and open the a link in a new tab

import requests,sys,webbrowser
from bs4 import BeautifulSoup

print("Googling...")
res = requests.get('http://google.com/search?q='+' '.join(sys.argv[1:]))
print(res.raise_for_status())

soup = BeautifulSoup(res.text,"html5lib")

linkElems = soup.select('div#main > div > div > div > a')

numOpen = min(5,len(linkElems))
for i in range(numOpen):
    webbrowser.open('http://google.com' + linkElems[i].get('href'))
