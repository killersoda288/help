import requests
from bs4 import BeautifulSoup
import xlsxwriter


def get_soup(url):
    r = requests.get(url)
    return BeautifulSoup(r.content, 'html.parser')



def write_xlsx(items, write_row):
    write_column = 0
    for item in items:
        worksheet.write(write_row, write_column, item)
        write_column += 1


workbook = xlsxwriter.Workbook('Results.xlsx')
worksheet = workbook.add_worksheet()

# user variables
while True:
    start_url = input('Url: ')
    if 'https://www.tripadvisor.com.sg/Hotels-' not in start_url:
        print(
            'Please enter a valid url. e.g https://www.tripadvisor.com.sg/Hotels-g255100-Melbourne_Victoria-Hotels.html')
    else:
        break

print('fetching page...')
soup = get_soup(start_url)

while True:
    min_rev_num = input('Min Reviews for property: ')
    if min_rev_num.isdigit():
        if int(min_rev_num) >= 0:
            min_rev_num = int(min_rev_num)
            break
    print('Please enter a valid number')

while True:
    print('Enter max number of low review number properties on a single page, from 0 to 30.')
    print('(Program will exit once this condition is fufilled)')
    num_rev_criteria = input('Input: ')
    if num_rev_criteria.isdigit():
        if 0 <= int(num_rev_criteria) <= 30:
            num_rev_criteria = int(num_rev_criteria)
            break

    print('Please enter a valid number')

while True:
    min_star_rating = input('Min star rating for property: ')
    if min_star_rating.isdigit():
        if 0 <= int(min_star_rating) <= 5:
            min_star_rating = float(min_star_rating)
            break

    print('Please enter a valid number')

while True:
    min_room_num = input('Min number of rooms: ')
    if min_room_num.isdigit():
        if int(min_room_num) >= 0:
            min_room_num = int(min_room_num)
            break
    print('Please enter a valid number')

# get num pages
while True:
    max_num_pages = int(soup.select_one('.pageNum.last.taLnk').text.strip())
    num_pages = input('Page to search until(1 to {}):'.format(str(max_num_pages)))
    if num_pages.isdigit():
        if 1 <= int(num_pages) <= max_num_pages:
            num_pages = int(num_pages)
            break
    print('Please enter a valid number')

write_row = 0
write_xlsx(['Property Details', 'Star Rating', 'Number of Rooms'], write_row)
page_url = start_url

print('Getting data...')
# get property data
for page_num in range(num_pages):
    print('On page {}'.format(str(page_num + 1)))
    low_review_count = 0
    soup = get_soup(page_url)
    if page_num != num_pages - 1:
        next_page = soup.select_one('.nav.next.taLnk.ui_button.primary')['href']
        page_url = 'https://www.tripadvisor.com.sg' + next_page
    else:
        pass
    rows = soup.select('.property_title.prominent')
    prop_urls = []
    for row in rows:
        prop_urls.append('https://www.tripadvisor.com.sg' + row['href'])
    for prop in prop_urls:
        soup = get_soup(prop)
        try:
            num_reviews = int(soup.select_one('.reviewCount').text.strip().split(' ')[0].replace(',', ''))
        except AttributeError:
            num_reviews = 0
        if num_reviews >= min_rev_num:
            try:
                property_name = soup.select_one('#HEADING').text.strip()
            except AttributeError:
                property_name = ' '

            try:
                star_rating_class = soup.select_one('.hotels-hotel-review-about-with-photos-layout-TextItem__textitem--3CMuR span')['class'][1]
                star_rating = float(star_rating_class[5] + '.' + star_rating_class[6])
            except TypeError:
                star_rating = 0

            num_rooms = 0
            extra_info = soup.select('.hotels-hotel-review-about-addendum-AddendumItem__content--28NoV')
            for data in extra_info:
                data = data.text.strip()
                if data.isdigit():
                    num_rooms = int(data)

            try:
                address = soup.select_one('.street-address').text.strip()+ ', ' + soup.select_one('.locality').text.strip() + soup.select_one('.country-name').text.strip()
            except AttributeError:
                address = ' '

            try:
                phone = soup.select_one('.is-hidden-mobile.detail').text.strip()
            except AttributeError:
                phone = ' '

            if star_rating >= min_star_rating or star_rating == 0:
                if num_rooms >= min_room_num or num_rooms == 0:
                    write_row += 1
                    write_xlsx([property_name + '\n' + address + '\nT: ' + phone, star_rating, num_rooms], write_row)
        else:
            low_review_count += 1

    if low_review_count >= num_rev_criteria:
        break
print('Done!')
workbook.close()

'''
Notes:
1) Will still get data if any of the parameters are missing,
   as long as the parameters that do exist meet the criteria
2) Current break criteria is for a page to have a certain number
   of entries with low reviews. This can be changed to suit needs.
3) This scraper relies on data from tripadvisor, which might not
   have much info on hotels in some destinations, like China.
4) Address and phone number are taken from tripadvisor as well.
   If you require time/dist to airport, and address/phone from google,
   must be done manually. Google maps does not allow these pages to
   be scraped.
5) Indiscriminate searching(no min num of reviews) will take 
   around 1min per page
ToDo:
1) Replace try and excepts with something less problematic
2) include vba script for formatting
3) Get inputs  through tkinter
4) Work on efficiency*
'''
