'use strict'

const Excel = require('exceljs')
const Moment = require('moment')
const Funcy = require('funcy')
const Combinatorics = require('./combinatorics.js')

const targetFile = "enseikun.xlsx"
const workbook = new Excel.Workbook()

workbook.xlsx.readFile(targetFile).then(function () {

	//入力
	const number_of_fleets = 3 //組み合わせる艦隊の数
	//const require_enseiID = [2,null,null]//含めたい遠征のID(3つまで、無い場合null)
	// const r = number_of_fleets - require_enseiID.filter(function(val){return val != null}).length
	//単発で出す遠征の最高時間
	const maxtime = {
		h: 8, //○時間
		m: 0, //○分

		init: function(){
			this.str = Moment("00:" + this.h + ":" + this.m, "HH:mm:ss").format("HH:mm:ss")
			return this
		}
	}.init()

	////報酬データ読込み
	const s = workbook.getWorksheet(1)
	let quests = []

	//2行から40行まで
	for (let i = 3; i <= 40; ++i) {
		let quest = {
			id : s.getCell(i, 1).value,
			name : s.getCell(i, 2).value,
			time : s.getCell(i, 3).value,
			once : {
				n: s.getCell(i, 10).value,
				d: s.getCell(i, 11).value,
				k: s.getCell(i, 12).value,
				b: s.getCell(i, 13).value
			},
			oncebig : {
				n: s.getCell(i, 14).value,
				d: s.getCell(i, 15).value,
				k: s.getCell(i, 16).value,
				b: s.getCell(i, 17).value
			},
			hour : {
				n: s.getCell(i, 18).value,
				d: s.getCell(i, 19).value,
				k: s.getCell(i, 20).value,
				b: s.getCell(i, 21).value
			},
			hourbig : {
				n: s.getCell(i, 22).value,
				d: s.getCell(i, 23).value,
				k: s.getCell(i, 24).value,
				b: s.getCell(i, 25).value
			},
			fleet : s.getCell(i, 26).value != null ? s.getCell(i, 26).value + " " : "" 
						+ s.getCell(i, 27).value != null ? s.getCell(i, 27).value + " " : "" 
						+ s.getCell(i, 28).value != null ? s.getCell(i, 28).value + " " : "" 
						+ s.getCell(i, 29).value != null ? s.getCell(i, 29).value : "",
			consumption : {
				n: s.getCell(i, 30).value,
				d: s.getCell(i, 31).value
			},
			ratio_consumption : {
				n: s.getCell(i, 30).value,
				d: s.getCell(i, 31).value
			},
			rate_suc : s.getCell(i, 34).value,
			rate_bigsuc : s.getCell(i, 35).value,
		}
		quests.push(quest)
	}

	//準備
	const number_of_quests = quests.length
	const ids = quests.map(function(val){
		return val.id
	})
	// 報酬情報を遠征idから取得
	const quest = function(id){	
		return quests.filter(function(elm){
			return elm.id == id
		}).pop()　//pop()は破壊的だが、filter()によって非破壊性を維持している
	}
	// 遠征の組合せ
	let allcmb = {
		cmb: Combinatorics.bigCombination(ids,number_of_fleets),
		patterns:[],

		init: function(){
			let elm
			while(elm = this.cmb.next())
				this.patterns.push(elm)

			//map start
			this.patterns = this.patterns.map(function(elm){
				const quests = elm.map(function(current,index,array){
						return quest(current)
					  })
				return {
					combi: elm ,
					once: {
						n: quests.reduce(function(x,y){
											return x + y.once.n * y.rate_suc + y.oncebig.n * y.rate_bigsuc
										}, 0),
						d: quests.reduce(function(x,y){
											return x + y.once.d * y.rate_suc + y.oncebig.d * y.rate_bigsuc
										}, 0),
						k: quests.reduce(function(x,y){
											return x + y.once.k * y.rate_suc + y.oncebig.k * y.rate_bigsuc
										}, 0),
						b: quests.reduce(function(x,y){
											return x + y.once.b * y.rate_suc + y.oncebig.b * y.rate_bigsuc
										}, 0),
						init: function(){
							this.sum = this.n + this.d + this.k + this.b
							this.nd = this.n + this.d
							return this
						}

					}.init(),

					hour: {
						n: quests.reduce(function(x,y){
											return x + y.hour.n * y.rate_suc + y.hourbig.n * y.rate_bigsuc
										}, 0),
						d: quests.reduce(function(x,y){
											return x + y.hour.d * y.rate_suc + y.hourbig.d * y.rate_bigsuc
										}, 0),
						k: quests.reduce(function(x,y){
											return x + y.hour.k * y.rate_suc + y.hourbig.k * y.rate_bigsuc
										}, 0),
						b: quests.reduce(function(x,y){
											return x + y.hour.b * y.rate_suc + y.hourbig.b * y.rate_bigsuc
										}, 0),
						init: function(){
							this.sum = this.n + this.d + this.k + this.b
							this.nd = this.n + this.d
							return this
						}

					}.init(),

					maxtime: quests.map(function(x){
											return x.time
										}).sort(function(a,b){
											return a < b
										}).shift(),
					
					cut: function(num){
						return Math.floor(num*100,2)/100
					},

					show: function(){
						return str = ""
						+ "▶ " + this.combi.join('-') + "\n" 
						+ "▶ " + quests.map(function(x){return x.name}).join(', ') + "\n"
						+ "▶ " + quests.map(function(x){return x.time}).join(', ') + "\n"
						+ "[単発]: <SUM合計> " + this.cut(this.once.sum)
						+ " <ND燃料+弾薬> " + this.cut(this.once.nd) + "\n"
						+ "[時給]: <SUM合計> " + this.cut(this.hour.sum)
						+ " <ND燃料+弾薬> " + this.cut(this.hour.nd)　+ "\n"
					}
				}
			})
			//map end
			return this
		},
		rank_once_nd: function(maxtime){
			return this.patterns.filter(function(a){
				return a.maxtime < maxtime.str
			}).sort(function(a,b){
				return b.once.nd - a.once.nd
			})			
		},	
		rank_once_sum: function(maxtime){
			return this.patterns.filter(function(a){
				return a.maxtime < maxtime.str
			}).sort(function(a,b){
				return b.once.sum - a.once.sum
			})			
		},
		rank_hour_nd: function(){
			return this.patterns.sort(function(a,b){
				return b.hour.nd - a.hour.nd
			})
		},
		rank_hour_sum: function(){
			return this.patterns.sort(function(a,b){
				return b.hour.sum - a.hour.sum
			})
		}
	}.init()

	////副作用
	console.log("指定最大遠征時間: " + maxtime.str + "\n")
	allcmb.rank_once_nd(maxtime).splice(1,1).forEach(function(val,index){
		console.log("●単発ND1位")
		console.log(val.show())
	})
	allcmb.rank_once_sum(maxtime).splice(1,1).forEach(function(val,index){
		console.log("●単発合計1位")
		console.log(val.show())
	})
	allcmb.rank_hour_nd().splice(1,1).forEach(function(val,index){
		console.log("●時給ND1位")
		console.log(val.show())
	})
	allcmb.rank_hour_sum().splice(1,1).forEach(function(val,index){
		console.log("●時給合計1位")
		console.log(val.show())
	})

})