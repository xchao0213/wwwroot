/**
 * jQuery jslides 1.1.0
 *
 * http://www.cactussoft.cn
 *
 * Copyright (c) 2009 - 2013 Jerry
 *
 * Dual licensed under the MIT and GPL licenses:
 *   http://www.opensource.org/licenses/mit-license.php
 *   http://www.gnu.org/licenses/gpl.html
 */
$(function(){
	var numpic = $('#slides li').size()-1;
	var nownow = 0;
	var inout = 0;
	var TT = 0;
	var SPEED = 5000;

	$('#slides li').eq(0).siblings('li').css({'display':'none'});
	var ulstart = '<ul id="pagination">',
		ulcontent = '',
		ulend = '</ul>';
	ADDLI();
	var pagination = $('#pagination li');
	var paginationwidth = $('#pagination').width();
	$('#pagination').css('margin-left',(-paginationwidth/2))
	pagination.eq(0).addClass('current')
	function ADDLI(){
		//var lilicount = numpic + 1;
		for(var i = 0; i <= numpic; i++){
			ulcontent += '<li>' + '<a href="javascript:void(0)">' + (i+1)+ '</a>' + '</li>';
		}
		
		$('#slides').after(ulstart + ulcontent + ulend);	
	}
	pagination.on('click',DOTCHANGE)
	function DOTCHANGE(){
		var changenow = $(this).index();
		$('#slides li').eq(nownow).css('z-index','9');
		$('#slides li').eq(changenow).css({'z-index':'8'}).show();
		pagination.eq(changenow).addClass('current').siblings('li').removeClass('current');
		$('#slides li').eq(nownow).fadeOut(400,function(){$('#slides li').eq(changenow).fadeIn(500);});
		nownow = changenow;
	}
	pagination.mouseenter(function(){
		inout = 1;
	})
	pagination.mouseleave(function(){
		inout = 0;
	})
	function GOGO(){
		var NN = nownow+1;
		if( inout == 1 ){
			} else {
			if(nownow < numpic){
			$('#slides li').eq(nownow).css('z-index','9');
			$('#slides li').eq(NN).css({'z-index':'8'}).show();
			pagination.eq(NN).addClass('current').siblings('li').removeClass('current');
			$('#slides li').eq(nownow).fadeOut(400,function(){$('#slides li').eq(NN).fadeIn(500);});
			nownow += 1;
		}else{
			NN = 0;
			$('#slides li').eq(nownow).css('z-index','9');
			$('#slides li').eq(NN).stop(true,true).css({'z-index':'8'}).show();
			$('#slides li').eq(nownow).fadeOut(400,function(){$('#slides li').eq(0).fadeIn(500);});
			pagination.eq(NN).addClass('current').siblings('li').removeClass('current');
			nownow=0;
			}
		}
		TT = setTimeout(GOGO, SPEED);
	}
	
	TT = setTimeout(GOGO, SPEED); 

})

$(function(){	
	
	$(".in_al_main li span").mouseover(function(){
		$(this).addClass("box").siblings("span").removeClass("box");
	}).mouseout(function(){
		$(this).removeClass("box").siblings("span");
	})
	

	var 
		index=0;
		Swidth=990;
		timer= null ;	
		function NextPage()
		{	
			if(index>shou)
			{
				index=0;
			}
			$(".in_al_main").stop(true, false).animate({left: -index*Swidth+"px"},600)		
		}
		
		function PrevPage()
		{	
			if(index<0)
			{
				index=shou;
			}
			$(".in_al_main").stop(true, false).animate({left: -index*Swidth+"px"},600)		
		}
		
		//下一页
		$(".next img").click(function(){
			 index++ ;
			 NextPage();
		});
		//上一页
		$(".prev img").click(function(){
			 index-- ;
			 PrevPage();
		});
		//自动滚动
		var timer = setInterval(function(){
				index++ ;
				NextPage();
			},4000);
			
		$(".next img , .in_al_main , .prev img").mouseover(function(){
			clearInterval(timer);
		});
		$(".next img , .in_al_main , .prev img").mouseleave(function(){
			timer = setInterval(function(){
				index++ ;
				NextPage();
			},4000);	
		});
			
})//建站套餐
