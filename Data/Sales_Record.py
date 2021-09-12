import pandas as pd
import matplotlib.pyplot as plt

####### Loading Data/File #######
d_f = pd.read_excel('100 Sales Records(Xlse).xlsx')
#print(d_f.head(10))

## Getting the headers 
headers = d_f.columns
print(headers)


####### Working with the loaded Data  #######

## Q1: Which columns(Country) has the highest(most) sales

## Getting the total cost price and Adding it to the data frame
d_f['Total Cost Price'] = d_f['Units Sold'] * d_f['Unit Cost']

## Getting the total cost price and Adding it to the data frame
d_f['Total Price'] = d_f['Units Sold'] * d_f['Unit Price']

## Getting the profit and adding it to the data frame
d_f['Profit'] = d_f['Total Price'] - d_f['Total Cost Price']


## Getting the country with the most sales
country_HighestSales = d_f.groupby('Country').max()[['Unit Cost','Units Sold', 'Unit Price', 'Total Cost Price', 'Total Price', 'Profit']]
# print(country_HighestSales)
country_HighestSales.to_excel('Country With The Most Sales.xlsx')


## Q2: Which columns(Month) has the highest sales
## Adding the month column and Getting the highest sales
"""Converting the column to datetime format and adding Month column"""

d_f['Ship Date'] = pd.to_datetime(d_f['Ship Date'])
d_f['Order Date'] = pd.to_datetime(d_f['Order Date'])
d_f['Month'] = d_f['Order Date'].dt.strftime('%m')
#print(d_f.head(10))
#print(d_f.iloc[0])
#print(d_f.iloc[2])

month_HighestSales = d_f.groupby('Month').sum()[['Units Sold', 'Total Cost Price', 'Profit']]
#print(month_HighestSales)
month_HighestSales.to_excel('The Month with the most sales.xlsx')


## Q3: What is the most effective sales channel
sale_cha = d_f.groupby('Sales Channel').sum()[['Units Sold', 'Unit Cost' ,'Total Cost Price', 'Unit Price', 'Total Price', 'Profit']]
#print(sale_cha)
sale_cha.to_excel('The Most Effective Sales Channel.xlsx')


## Q4: Which product is the most sold
item_sale = d_f.groupby('Item Type').sum()[['Units Sold', 'Unit Cost' ,'Total Cost Price', 'Unit Price', 'Total Price', 'Profit']]
#print(item_sale)
item_sale.to_excel('The Most Sold Product.xlsx')


## Q5: Which country and region buys the most sold product
country_item_sale = d_f[d_f['Item Type'] == 'Cosmetics'][['Region', 'Country', 'Units Sold', 'Unit Cost' ,'Total Cost Price', 'Unit Price', 'Total Price', 'Profit']]
#print(country_item_sale)
cosmetics_country = country_item_sale.groupby('Country').sum()[['Units Sold', 'Unit Cost' ,'Total Cost Price', 'Unit Price', 'Total Price', 'Profit']]
#print(cosmetics_country)
cosmetics_country.to_excel('Countries that brought the most sold product.xlsx')

## Q6: Which country has the highest cost price
## Getting the country with highest Cost Price
totalCP = d_f.groupby('Country').sum()[['Units Sold', 'Total Cost Price', 'Profit']]
#print(totalCP)
totalCP.to_excel('Country and Cost Price.xlsx')

## Q7: Which columns(Country)/ (product) has the highest profit

HighestProfit_country = d_f.groupby('Country').max()[['Item Type', 'Total Price', 'Profit']]
#print(HighestProfit_country)
HighestProfit_country.to_excel('Country and Profit.xlsx')

HighestProfit_item =  d_f.groupby('Item Type').sum()[['Profit']]
#print(HighestProfit_item)
HighestProfit_item.to_excel('Products and their profits.xlsx')




### Visualization Of Data ###

plt.style.use('seaborn')

## V1: Which columns(Country) has the highest(most) sales
# countries = [country for country, df in d_f.groupby('Country')]
# plt.bar(countries,   country_HighestSales['Units Sold'], label='Country')
# plt.xticks(countries, rotation='vertical', size=8)
# plt.title('The Country With The Most Sales')
# plt.xlabel('Countries')
# plt.ylabel('Unit Sold')
# plt.grid(True, color='black')
# plt.show()


# ## V2: Which columns(Month) has the highest sales
# months = ['January', 'February', 'March', 'April', 'May', 
# 'June', 'July', 'August', 'September', 'October', 'November', 'December']
# plt.bar(months,   month_HighestSales['Units Sold'])
# plt.xticks(months)
# plt.title('The Month with the highest sales')
# plt.xlabel('Month')
# plt.ylabel('Unit Sold')
# plt.grid(True, color='black')
# plt.show()

## V3: What is the most effective sales channel
# chan = ['Offline', 'Online']
# plt.bar(chan, sale_cha['Units Sold'], width=0.5, align='center')
# plt.title('The Most Effective Sales Channel')
# plt.xticks(chan)
# plt.xlabel('Sales Channel')
# plt.ylabel('Unit Sold')
# plt.grid(True, color='black')
# plt.show()

## V4: Which product is the most sold
# items = [item for item, df in d_f.groupby('Item Type')]
# plt.bar(items,  item_sale['Units Sold'])
# plt.title('The Most Sold Product')
# plt.xticks(items, rotation='vertical', size=8)
# plt.xlabel('Products')
# plt.ylabel('Units Sold')
# plt.grid(True, color='black')
# plt.show()

## V5: Which country and region buys the most sold product
# countryItemSale = [con for con, df in cosmetics_country.groupby('Country')]
# plt.bar(countryItemSale,  cosmetics_country['Units Sold'])
# plt.title('Countries That Brought The Most Sold Product')
# plt.xticks(countryItemSale, rotation='vertical', size=8)
# plt.xlabel('Country')
# plt.ylabel('Units Brought')
# plt.grid(True, color='black')
# plt.show()


## V6: Which country has the highest cost price
# countryItemSale = [con for con, df in totalCP.groupby('Country')]
# plt.bar(countryItemSale,  totalCP['Total Cost Price'])
# plt.title('Country With Their Cost Price')
# plt.xticks(countryItemSale, rotation='vertical', size=8)
# plt.xlabel('Country')
# plt.ylabel('Total Cost Price (in Million)', rotation='vertical')
# plt.grid(True, color='black')
# plt.show()


## V7: Which country has the highest profit
# count_pro = [con for con, df in HighestProfit_country.groupby('Country')]
# plt.bar(count_pro,  HighestProfit_country['Profit'])
# plt.title('Country And Their Profit')
# plt.xticks(count_pro, rotation='vertical', size=8)
# plt.xlabel('Country')
# plt.ylabel('Profit', rotation='vertical')
# plt.grid(True, color='black')
# plt.show()


## V8: Which item has the highest profit
# count_pro = [con for con, df in HighestProfit_item.groupby('Item Type')]
# plt.plot(count_pro, HighestProfit_item['Profit'])
# plt.title('Product and Profit')
# plt.xticks(count_pro, rotation='vertical')
# plt.xlabel('Product')
# plt.ylabel('Profit')
# plt.grid(True, color='black')
# plt.show()
