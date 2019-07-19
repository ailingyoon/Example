/***********
 * name:   ajaxRequest
 * author: qiongmiaoer
 * date:   2015-2-13
 * desc:   ajax方法封装提取，及使用demo范例
 * **********/

/*
 * description:
 * 
 * ajax请求 区分同步 Synchronous、异步Asynchronous
 * params:
 *    url:请求url
 *    data:post请求data参数 *,
 *    callback:回调函数    
 * */
var AjaxRequest={};
//get 异步
AjaxRequest.get=function(url,callback){ 
	$.ajax({
		type: 'GET',
		url: url,
		//async:true,   //ajax请求默认'异步'
		success: function(result){
			callback(result);
		},
		contentType:"application/json"
	});
}
//get 同步
AjaxRequest.getAsync=function(url,callback){ 
	$.ajax({
		type: 'GET',
		url: url,
		async:false,
		success: function(result){
			callback(result);
		},
		contentType:"application/json"
	});
}

//post 异步
AjaxRequest.post=function(url,data,callback){
	$.ajax({
		  type: 'POST',
		  url: url,
		  data: JSON.stringify(data),
		  success: function(result){
			  callback(result);
		  },
		  dataType: "json",
		  contentType:"application/json"
		});
}
//post 同步
AjaxRequest.postAsync=function(url,data,callback){
	$.ajax({
		type: 'POST',
		url: url,
	    async:false, 
		data: JSON.stringify(data),
		success: function(result){
			callback(result);
		},
		dataType: "json",
		contentType:"application/json"
	});
}

//处理ajax 数据返回值
AjaxRequest.resultHandler=function(result){
	//success
	if(result.code==1){
		return true;
	}else{ //error
		//页面友好处理
		alert('error');
		return false;
	}
}

/***************demo***************/
//get 异步
var demo_GetFunc=function(){
	AjaxRequest.get('rest/weather/getWeatherList/20150123/',function(result){
		//result eg
		result={ 
				code:1,  //1 success 0 error
				messge:'ajax request success',
				data:{
					date:'2015-1-23',
					type:'Industry',
					list:[
					      {id:1,name:'industry_1'},{id:2,name:'industry_2'},{id:3,name:'industry_3'}
					]
				}
		};
		var status = AjaxRequest.resultHandler(result);
		if(!status){
			//页面当前模块处理方法
			
			return;
		}
		//result 数据分析与展示
	});
}
//get 同步
var demo_PostFunc2=function(){
	var data={id:1};
	AjaxRequest.getAsync('rest/weather/getWeatherInfoById/'+data,function(result){
		var status = AjaxRequest.resultHandler(result);
		//同上 
	});
}

//post 异步
var demo_PostFunc2=function(){
	var data={id:1};
	AjaxRequest.post('rest/weather/getWeatherList/20150123/',function(result){
		var status = AjaxRequest.resultHandler(result);
		//同上 
	});
}
//post 同步
var demo_PostFunc=function(){
	var data={id:1};
	AjaxRequest.postAsync('rest/weather/getWeatherInfoById/'+data,function(result){
		var status = AjaxRequest.resultHandler(result);
		//同上 
	});
}
