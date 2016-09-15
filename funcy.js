'use strict'

//normal ver
const input0 = 0
const input1 = 1
let output
switch(input0){
	case 0:
		output = "あ"
		break
	default:
		output = "うん"
		break
}
console.log(output)
switch(input1){
	case 0:
		output = "あ"
		break
	default:
		output = "うん"
		break
}
console.log(output)

//funcy ver
const fun = require('funcy')
const $ = fun.parameter;
const func = fun(
	[0,function(){return "あ"}],
	[$,function(){return "うん"}]
)
console.log(func(0))
console.log(func(1))