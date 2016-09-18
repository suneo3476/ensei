// this code runs standalone
// 入力:13行目　出力:135行目
(function(){
	'use strict'

	const Excel = require('exceljs')
	const Moment = require('moment')
	const Funcy = require('funcy')
	const _ = require('underscore')
	const Combinatorics = require('./combinatorics.js')

	const main = function () {
		//ユーザ入力
		const number_of_fleets = 3 //組み合わせる艦隊の数
		const require_enseiID = [null,null,null]//含めたい遠征のID(3つまでカンマ(,)区切りで入力、希望がない場合はnull)
								.splice(0,number_of_fleets).filter(function(x){return x!=null})
		const maxtime = {//単発で出す遠征の最高時間
			h: 10, //○時間
			m: 0, //○分
			init: function(){
				this.str = Moment("00:" + this.h + ":" + this.m, "HH:mm:ss").format("HH:mm:ss")
				return this
			}
		}.init()

		////報酬データ
		const quests = load_quests()

		// 遠征idによる各遠征情報の取得
		const quest = function(id){	
			return quests.filter(function(elm){
				return elm.id == id
			}).pop()　//pop()は破壊的だが、filter()によって非破壊性を維持している
		}

		// 遠征の組合せ
		const allcmb = {
			cmb: Combinatorics.bigCombination(　quests.map(function(val){return val.id})　,　number_of_fleets　),
			patterns:[],
			init: function(){
				this.patterns = function(cmb){
					const patterns = []
					for(let elm;elm = cmb.next();)
						patterns.push(elm)
					return patterns.map(function(elm){
						const quests = elm.map(function(current,index,array){
							return quest(current)
						})
						return {
							combi: elm ,
							once: {
								n: quests.reduce(function(x,y){return x + y.once.n * y.rate_suc + y.oncebig.n * y.rate_bigsuc}, 0),
								d: quests.reduce(function(x,y){return x + y.once.d * y.rate_suc + y.oncebig.d * y.rate_bigsuc}, 0),
								k: quests.reduce(function(x,y){return x + y.once.k * y.rate_suc + y.oncebig.k * y.rate_bigsuc}, 0),
								b: quests.reduce(function(x,y){return x + y.once.b * y.rate_suc + y.oncebig.b * y.rate_bigsuc}, 0),
								init: function(){
									this.sum = this.n + this.d + this.k + this.b
									this.nd = this.n + this.d
									return this
								}
							}.init(),
							hour: {
								n: quests.reduce(function(x,y){return x + y.hour.n * y.rate_suc + y.hourbig.n * y.rate_bigsuc}, 0),
								d: quests.reduce(function(x,y){return x + y.hour.d * y.rate_suc + y.hourbig.d * y.rate_bigsuc}, 0),
								k: quests.reduce(function(x,y){return x + y.hour.k * y.rate_suc + y.hourbig.k * y.rate_bigsuc}, 0),
								b: quests.reduce(function(x,y){return x + y.hour.b * y.rate_suc + y.hourbig.b * y.rate_bigsuc}, 0),
								init: function(){
									this.sum = this.n + this.d + this.k + this.b
									this.nd = this.n + this.d
									return this
								}
							}.init(),
							maxtime: quests.map(function(x){return x.time}).sort(function(a,b){return a < b}).shift(),
							cut: function(num){return Math.floor(num*100,2)/100},
							init: function(){
								const cut = this.cut
								this.show = {
									combi: this.combi.join('-'),
									name: quests.map(function(x){return x.name}).join(', '),
									time: quests.map(function(x){return Moment(x.time,"HH:mm:ss").format("mm:ss")}).join(', '),
									oncen: '燃料 ' + cut(this.once.n) + "(弾薬 " + cut(this.once.d) + ")",
									onced: '弾薬 ' + cut(this.once.d) + "(燃料 " + cut(this.once.n) + ")",
									oncek: '鋼材 ' + cut(this.once.k),
									onceb: 'ボーキ ' + cut(this.once.b),
									hourn: '燃料 ' + cut(this.hour.n) + "(弾薬 " + cut(this.hour.d) + ")",
									hourd: '弾薬 ' + cut(this.hour.d) + "(燃料 " + cut(this.hour.n) + ")",
									hourk: '鋼材 ' + cut(this.hour.k),
									hourb: 'ボーキ ' + cut(this.hour.b),
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
				}(this.cmb)
				return this
			},
			sort: {
				once_nd: function(x,y){return y.once.nd - x.once.nd},
				hour_nd: function(x,y){return y.hour.nd - x.hour.nd},
				once_sum: function(x,y){return y.once.sum - x.once.sum},
				hour_sum: function(x,y){return y.hour.sum - x.hour.sum},
				once_d: function(x,y){return y.once.d - x.once.d},
				hour_d: function(x,y){return y.hour.d - x.hour.d},
				once_n: function(x,y){return y.once.n - x.once.n},
				hour_n: function(x,y){return y.hour.n - x.hour.d},
				once_k: function(x,y){return y.once.k - x.once.k},
				hour_k: function(x,y){return y.hour.k - x.hour.k},
				once_b: function(x,y){return y.once.b - x.once.b},
				hour_b: function(x,y){return y.hour.b - x.hour.b},
			},
			calcrank: function(sort,maxtime){
				return this.patterns.filter(function(x){
					return x.combi.some(function(x){
						return require_enseiID.length != 0 ? _.contains(require_enseiID, x) : true //含めたい遠征IDを含む遠征パターンを抽出する
					}) && (x.maxtime <= maxtime.str) //遠征の最大時間を超えない遠征パターンを抽出する
				}).sort(sort)
			},
			viewrank: function(sort,maxtime){
				_.each(this.calcrank(sort,maxtime).splice(0,1),function(x){
					const y = x.show
					const str = "[" + y.combi + "] " + y.time + "\n" + y.name + "\n" + y.onced + "\n"
					console.log(str)
				})
			}
		}.init()

		//出力
		console.log("艦隊数: " + number_of_fleets)
		console.log("含める遠征のID: "　+ require_enseiID.length != 0 ? require_enseiID : "なし")
		console.log("単発最大遠征時間: " + maxtime.str + "\n")
		//ソート方法は106行目～から選ぶ
		console.log("単発-弾薬-1位")
		allcmb.viewrank(allcmb.sort.once_d,maxtime)

		console.log("時給-弾薬-1位")
		allcmb.viewrank(allcmb.sort.hour_d,maxtime)

	}

	//報酬データの読込み
	const load_quests = function(){
		const s = workbook.getWorksheet(1)
		let quests = []
		for (let i = 3; i <= 40; ++i) {
			const quest = {
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
		return quests
	}

	const workbook = new Excel.Workbook()
	const targetFile = "enseikun.xlsx"
	workbook.xlsx.readFile(targetFile).then(main)

})()