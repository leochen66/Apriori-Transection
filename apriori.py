import pandas as pd

FILE_NAME = 'transection.xlsx'
RESULT_NAME = 'result.xlsx'

# min_support = 0.004
# min_confident = 0.004
# min_support_count = len(db.index) * min_support
min_support_count = 1000

# input: transection id
# outpit： list that contain items in sepecific trasection
def get_single_transection(n):
	items = []
	items.append(db.loc[n]['品項 1'])
	items.append(db.loc[n]['品項 2'])
	items.append(db.loc[n]['品項 3'])
	items.append(db.loc[n]['品項 4'])
	items.append(db.loc[n]['品項 5'])
	return items

# input: transection id
# output: list that contain items' number in sepecific trasection
def get_item_number(n):
	items_num = []
	items_num.append(db.loc[n]['品項 1 數量'])
	items_num.append(db.loc[n]['品項 2 數量'])
	items_num.append(db.loc[n]['品項 3 數量'])
	items_num.append(db.loc[n]['品項 4 數量'])
	items_num.append(db.loc[n]['品項 5 數量'])
	return items_num

# input: two lists(subli is the sublist of li)
# output: the index of sublist elements in li
def find_sub(li, subli):
	location = []
	for i in subli:
		for j in li:
			if j == i:
				location.append(li.index(j))
				break
	return location

# read data from exel
db = pd.read_excel(FILE_NAME, index_col='訂單號碼')

# Aprioi: find one-frquent-itemset
# calculate the count of each item, and delete items whose count is lower than min_support_count
k = 1
result = {'組合':[], '組合機率':[], '數量':[], '數量機率':[]}
itemsets = []
frequency = []
one_itemset = {}
for i in db.index:
	for item in get_single_transection(i):
		if item not in one_itemset:
			one_itemset[item] = 1
		else:
			one_itemset[item] = one_itemset[item] +1
print(one_itemset)
frequent_itemset = []
for item in one_itemset:
	if one_itemset[item] >= min_support_count:
		frequent_itemset.append([item])
# print(frequent_itemset)


# Loop until no sub-itemset genarated
while True:

	# Apriori: Genarate k+1 itemset
	# Combine rule: different on index K-1
	new_itemset = []
	if k ==1:
		for i in range(len(frequent_itemset)):
			for j in range(i+1, len(frequent_itemset)):
				temp = [frequent_itemset[i][0], frequent_itemset[j][0]]
				new_itemset.append(temp)
		#print('new_itemset: ', new_itemset)
	else:
		done_list = []
		for i in range(len(frequent_itemset)):
			for j in range(i+1, len(frequent_itemset)):
				if (i not in done_list) and (j not in done_list) and frequent_itemset[i][k-1] != frequent_itemset[j][k-1] and frequent_itemset[i][0:k-1] == frequent_itemset[j][0:k-1]:
					new_itemset.append(frequent_itemset[i][0:k-1] + [frequent_itemset[i][k-1]] + [frequent_itemset[j][k-1]])
					done_list.append(i)
					done_list.append(j)
					break
		#print('new_itemset: ', new_itemset)

	# If no new sub-itemset genarated, end loop
	if len(new_itemset) ==0:
		break

	# Apriori: delete itemset that is not frequent
	prune_itemset = []
	for itemset in new_itemset:
		for i in range(len(itemset)):
			temp = itemset.pop(i)
			if itemset in frequent_itemset:
				itemset.insert(i, temp)
				prune_itemset.append(itemset)
				break
			itemset.insert(i, temp)
	#print('prune_itemset: ', prune_itemset)

	# Apriori: calculate the probability and delete itemset whose probability lower than min_support
	frequent_itemset = []  # refresh frequent_itemset

	print('Count of Scans: ', len(prune_itemset))
	print('Time Speculation: ', len(prune_itemset)*20,'秒')
	t = 0

	for itemset in prune_itemset:
		t = t+1
		print(str(t) + 'Scan Started!')

		count = 0
		set_itemset = set(itemset)
		for i in db.index: # scan database
			if set_itemset.issubset(get_single_transection(i)):
				count = count +1
		if count >= min_support_count:
			frequent_itemset.append(itemset)
			itemsets.append(itemset)
			frequency.append(count / len(db.index))		
	# Go to next round
	k = k+1

##########################################################################

# calculate probability
count_itemset = {}
for item in itemsets:
	temp_dic = {}
	set_item = set(item)
	for i in db.index:
		if set_item.issubset(get_single_transection(i)):
			item_num = get_item_number(i)
			location = find_sub(get_single_transection(i), item)
			num_array = []
			for l in location:
				num_array.append(item_num[l])
			num_array = str(num_array)
			if num_array not in temp_dic:
				temp_dic[num_array] = 1
			else:
				temp_dic[num_array] = temp_dic[num_array] +1
	count_itemset[str(item)] = temp_dic
# print(count_itemset)


# write the result to exel
for item in itemsets:
	numbers = count_itemset[str(item)]
	for num in numbers:
		result['數量'].append(num)
		result['數量機率'].append(numbers[num] / len(db.index))
	result['組合'].append(item)
	result['組合機率'].append(frequency[itemsets.index(item)])
	while len(result['數量']) > len(result['組合']):
		result['組合'].append('')
		result['組合機率'].append('')
result = pd.DataFrame(result)
result.to_excel(RESULT_NAME)