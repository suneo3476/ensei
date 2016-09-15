const array = [1,3,5]
const array2 = [3,5,7]
const array3 = [null,"a",""]
const sum = [1,3,5].reduce(function(pre, current){
	console.log(pre, current)
	return pre + current*2
},initialvalue=0)
console.log(sum)