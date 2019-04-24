rechargeKeys = ["coreMultl" , 
	"coreFacilityFee" , 
	"assocMult" , 
	"assocFacilityFee" , 
	"regMult" , 
	"regFacilityFee" , 
	"costHDP" , 
	"costSDP" , 
	"costMosqTip" , 
	"costMosquitoTime" , 
	"costDragonflyTime" , 
	"costRockImagerTime" , 
	"costMXOne" , 
	"costHDSConsume" , 
	"costSDSConsume" , 
	"costDFlyReservoir" , 
	"costDFlyTip" , ]

rechargeTuples = [('Core use multiplier', 'Price'),
	('Core facility fee','Price'),
	('Assoc use multiplier','Price'),
	('Assoc facility fee','Price'),
	('Regular use multiplier','Price'),
	('Regular facility fee','Price'),
	('96-well Greiner Hanging drop plate','Price/Qty'), 
	('MRC2 Sitting drop plate','Price/Qty'), 
	('Spool of mosquito tips 9 mm pitch','Price/Qty'),
	('Mosquito time','Price'),
	('Dragonfly time','Price'),
	('RockImager time','Price'),
	('MXone pin arrays','Price/Qty'),
	('HD Consumables','Price'), 
	('SD Consumables','Price'), 
	('Reservoir dragonfly','Price/Qty'),
	('Pack of 100 dragonfly tips','Price/Qty'),]

def getRechargeConsts(df, rowHeader, keys, tuples):
	consts_dict = {}

	for k in range(len(keys)):
		row, col = tuples[k]
		consts[keys[k]] = currencyToFloat(getCellByRowCol(df, 
			rowHeader, row, col))
	return consts_dict