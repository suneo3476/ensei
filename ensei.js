'use strict'

const Excel = require('exceljs')
const Moment = require('moment')
const Funcy = require('funcy')
const _ = require('underscore')
const Combinatorics = require('./combinatorics.js')

const targetFile = "enseikun.xlsx"
const workbook = new Excel.Workbook()

workbook.xlsx.readFile(targetFile).then(function () {
	//入力
	const number_of_fleets = 3 //組み合わせる艦隊の数
	const require_enseiID = [2,null,null]//含めたい遠征のID(3つまでカンマ(,)区切りで入力、希望がない場合はnull)
							.splice(0,number_of_fleets).filter(function(x){return x!=null})
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

					init: function(){
						const cut = this.cut
						this.show = {
							combi: this.combi.join('-'),
							name: quests.map(function(x){return x.name}).join(', '),
							time: quests.map(function(x){return Moment(x.time,"HH:mm:ss").format("mm:ss")}).join(', '),

							oncend: '燃弾 ' + cut(this.once.nd),
							oncendeach: '燃 ' + cut(this.once.n) + ' 弾 ' + cut(this.once.d),

							oncesum: '合計' + cut(this.once.sum),
							oncesumeach: '燃 ' + cut(this.once.n) + ' 弾 ' + cut(this.once.d) + ' 鋼 ' + cut(this.once.k) + ' ボ ' + cut(this.once.b),
										
							hournd: '燃弾 ' + cut(this.hour.nd),
							hourndeach: '燃 ' + cut(this.hour.n) + ' 弾 ' + cut(this.hour.d),

							hoursum: '合計' + cut(this.hour.sum),
							hoursumeach: '燃 ' + cut(this.hour.n) + ' 弾 ' + cut(this.hour.d) + ' 鋼 ' + cut(this.hour.k) + ' ボ ' + cut(this.hour.b)
						}
						return this
					}
				}.init()
			})
			//map end
			return this
		},
		rank_once_nd: function(maxtime){
			return this.patterns.filter(function(x){
				return x.combi.some(function(x){
					return _.contains(require_enseiID, x)
				}) && (x.maxtime <= maxtime.str)
			}).sort(function(x,y){
				return y.once.nd - x.once.nd
			})			
		},
		rank_once_sum: function(maxtime){
			return this.patterns.filter(function(x){
				return x.combi.some(function(x){
					return _.contains(require_enseiID, x)
				}) && (x.maxtime <= maxtime.str)
			}).sort(function(x,y){
				return y.once.sum - x.once.sum
			})			
		},
		rank_hour_nd: function(){
			return this.patterns.filter(function(x){
				return x.combi.some(function(x){
					return _.contains(require_enseiID, x)
				})
			}).sort(function(x,y){
				return y.hour.nd - x.hour.nd
			})
		},
		rank_hour_sum: function(){
			return this.patterns.filter(function(x){
				return x.combi.some(function(x){
					return _.contains(require_enseiID, x)
				})
			}).sort(function(x,y){
				return y.hour.sum - x.hour.sum
			})
		}
	}.init()

	////副作用
	console.log("単発最大遠征時間: " + maxtime.str + "\n")
	allcmb.rank_once_nd(maxtime).splice(0,1).forEach(function(val,index){
		const x = val.show
		const str = ""
		+ "●単発燃弾1位" + " "
		+ x.combi + " " + x.time + "\n" + x.name + "\n"
		+ x.oncend + " " + x.oncendeach + "\n"
		console.log(str)
	})
	allcmb.rank_once_sum(maxtime).splice(0,1).forEach(function(val,index){
		const x = val.show
		const str = ""
		+ "●単発合計1位" + " "
		+ x.combi + " " + x.time + "\n" + x.name + "\n"
		+ x.oncesum + " " + x.oncesumeach + "\n"
		console.log(str)
	})
	allcmb.rank_hour_nd().splice(0,1).forEach(function(val,index){
		const x = val.show
		const str = ""
		+ "●時給燃弾1位" + " "
		+ x.combi + " " + x.time + "\n" + x.name + "\n"
		+ x.hournd + " " + x.hourndeach + "\n"
		console.log(str)
	})
	allcmb.rank_hour_sum().splice(0,1).forEach(function(val,index){
		const x = val.show
		const str = ""
		+ "●時給合計1位" + " "
		+ x.combi + " " + x.time + "\n" + x.name + "\n"
		+ x.hoursum + " " + x.hoursumeach + "\n"
		console.log(str)
	})

})