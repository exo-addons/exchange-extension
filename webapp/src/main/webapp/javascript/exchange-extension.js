function addExchangeButton() {
	var separatorLineElement = $( ".UICalendarPortlet .calendarWorkingWorkspace .uiActionBar .btnRight .separatorLine" );
	if(!separatorLineElement || separatorLineElement.length == 0) {
		return;
	}
	separatorLineElement.before("<a href='#' class='ExchangeSettingsButton pull-right'><img src='/exchange-resources/skin/images/exchange-disabled.png' width='24px' height='24px'/></a>");
	$('.ExchangeSettingsButton').click(function(e) {
		$(".ExchangeEditSettingsButton").removeAttr("disabled");
		
		$('.ExchangeSettingsWindow .ExchangeSettingsContent').html("<div class='ExchangeSettingsLoading'>Loading...</div>");
    	$.getJSON("/portal/rest/exchange/calendars", function(data){
    		$('.ExchangeSettingsWindow .ExchangeSettingsContent').html("");
        	if(!data || data.length == 0) {
    			$('.ExchangeSettingsButton img').attr('src', '/exchange-resources/skin/images/exchange-disabled.png');
    			$('.ExchangeSettingsWindow .ExchangeSettingsContent').html('<div class="ExchangeSettingsError">User seems not connected to Exchange</div>');
    		} else {
    			$('.ExchangeSettingsButton img').attr('src', '/exchange-resources/skin/images/exchange.png');
	    	    $.each(data, function(i,item){
	    	    	$('.ExchangeSettingsWindow .ExchangeSettingsContent').append(""+item.name+"<input type='checkbox' "+(item.synchronizedFolder?"checked":"")+" name='"+item.name+"' value='"+item.id+"' /><BR/>");
	    	    });
	        	$('.ExchangeSettingsWindow input[type="checkbox"]').click(function(){
	        	    if($(this).is(':checked')){
	        	    	$.get("/portal/rest/exchange/sync?"+$.param({folderId : $(this).val()}));
	        	    } else {
	        	    	$.get("/portal/rest/exchange/unsync?"+$.param({folderId : $(this).val()}));
	        	    }
	        	});
    		}
    	});
    	$('.ExchangeSettingsWindow').css('top', (separatorLineElement.position().top + 25) + 'px');
    	$('.ExchangeSettingsWindow').css('right', ($(window).width() - separatorLineElement.position().left - 37) + 'px');

		$(".ExchangeEditSettingsPanel").hide();
		$(".ExchangeSettingsContent").show();
    	$('.ExchangeSettingsMask').show();
	    $('.ExchangeSettingsWindow').show();
	});

	$.getJSON("/portal/rest/exchange/calendars", function(data){
    	if(!data || data.length == 0) {
			$('.ExchangeSettingsButton img').attr('src', '/exchange-resources/skin/images/exchange-disabled.png');
		} else {
			$('.ExchangeSettingsButton img').attr('src', '/exchange-resources/skin/images/exchange.png');
		}
	});

	if(!$('.ExchangeSettingsWindow') || $('.ExchangeSettingsWindow').length == 0) {
		$("body").append("<div class='ExchangeSettingsWindow' />");
	}
	$('.ExchangeSettingsWindow').hide();
	$('.ExchangeSettingsWindow').html("<div class='ExchangeSettingsTitle'><h6>Exchange Calendars</h6><button type='button' class='btn btn-primary ExchangeEditSettingsButton'>Edit settings</button></div><div class='ExchangeEditSettingsPanel'><div class='ExchangeEditSettingsTitle'></div><div class='ExchangeEditSettingsContent'></div><div class='ExchangeEditSettingsButtons'><button type='button' class='btn btn-primary ExchangeEditSettingsSaveButton'>Save</button><button type='button' class='btn ExchangeEditSettingsCancelButton'>Cancel</button></div></div><div class='ExchangeSettingsContent'></div>");
	$(".ExchangeEditSettingsPanel").hide();
	$(".ExchangeEditSettingsContent").html("<label for='serverName'>URL</label><input type='text' id='serverName' name='serverName' placeholder='http://server/EWS/Exchange.asmx'><br/>");
	$(".ExchangeEditSettingsContent").append("<label for='domainName'>Domain</label><input type='text' id='domainName' name='domainName' placeholder='Exchange Domain Name'><br/>");
	$(".ExchangeEditSettingsContent").append("<label for='username'>Username</label><input type='text' id='username' name='username' placeholder='Required'><br/>");
	$(".ExchangeEditSettingsContent").append("<label for='password'>Password</label><input type='password' id='password' name='password' placeholder='Required'><br/>");

	if(!$('.ExchangeSettingsMask') || $('.ExchangeSettingsMask').length == 0) {
		$("body").append("<div class='ExchangeSettingsMask' />");
	}
    $('.ExchangeSettingsMask').hide();
	$('.ExchangeSettingsMask').click(function(e) {
		if (e.target.id == 'ExchangeSettingsMask') {
			return true;
		} else {
		    $('.ExchangeSettingsMask').hide();
			$('.ExchangeSettingsWindow').hide();
		}
	});
	
	$(".ExchangeEditSettingsButton").click(function(e) {
		$(".ExchangeEditSettingsContent #username").val("");
		$(".ExchangeEditSettingsContent #password").val("");
		$(".ExchangeEditSettingsContent #domainName").val("");
		$(".ExchangeEditSettingsContent #serverName").val("");
		
    	$(".ExchangeEditSettingsContent input").removeAttr("style");
    	$(".ExchangeEditSettingsContent label").removeAttr("style");
		
    	$.getJSON("/portal/rest/exchange/settings", function(data){
        	if(!data || data.length == 0) {
        		$('.ExchangeSettingsWindow .ExchangeSettingsContent').html('<div class="ExchangeSettingsError">Error getting settings from eXo Server.</div>');
        		return;
        	}
        	if(data.serverName) {
        		$(".ExchangeEditSettingsContent #serverName").val(data.serverName);
        	} else {
        	}
        	if(data.domainName) {
        		$(".ExchangeEditSettingsContent #domainName").val(data.domainName);
        	} else {
        		$(".ExchangeEditSettingsContent #domainName").val("");
        	}
    		if(data.username) {
        		$(".ExchangeEditSettingsContent #username").val(data.username);
        	} else {
        		$(".ExchangeEditSettingsContent #username").val("");
        	}
    	});
		$(".ExchangeEditSettingsButton").attr("disabled", "true");
		$(".ExchangeSettingsContent").hide();
		$(".ExchangeEditSettingsPanel").show();
	});
	$(".ExchangeEditSettingsCancelButton").click(function(e) {
		$(".ExchangeSettingsWindow").hide();
		$('.ExchangeSettingsButton').click();
	});
	$(".ExchangeEditSettingsSaveButton").click(function(e) {
		var exchangeSettingsNOK = false;
    	if(!$(".ExchangeEditSettingsContent #serverName").val()) {
    		$(".ExchangeEditSettingsContent #serverName").css("border-color", "red");
    		$(".ExchangeEditSettingsContent label[for='serverName']").css("color", "red");
    		exchangeSettingsNOK = true
    	} else {
        	$(".ExchangeEditSettingsContent #serverName").removeAttr("style");
        	$(".ExchangeEditSettingsContent label[for='serverName']").removeAttr("style");
    	}
    	if(!$(".ExchangeEditSettingsContent #domainName").val()) {
    		$(".ExchangeEditSettingsContent #domainName").css("border-color", "red");
    		$(".ExchangeEditSettingsContent label[for='domainName']").css("color", "red");
    		exchangeSettingsNOK = true
    	} else {
        	$(".ExchangeEditSettingsContent #domainName").removeAttr("style");
        	$(".ExchangeEditSettingsContent label[for='domainName']").removeAttr("style");
    	}
		if(!$(".ExchangeEditSettingsContent #username").val()) {
    		$(".ExchangeEditSettingsContent #username").css("border-color", "red");
    		$(".ExchangeEditSettingsContent label[for='username']").css("color", "red");
    		exchangeSettingsNOK = true
    	} else {
        	$(".ExchangeEditSettingsContent #username").removeAttr("style");
        	$(".ExchangeEditSettingsContent label[for='username']").removeAttr("style");
    	}
		if(!$(".ExchangeEditSettingsContent #password").val()) {
    		$(".ExchangeEditSettingsContent #password").css("border-color", "red");
    		$(".ExchangeEditSettingsContent label[for='password']").css("color", "red");
    		exchangeSettingsNOK = true
    	} else {
        	$(".ExchangeEditSettingsContent #password").removeAttr("style");
        	$(".ExchangeEditSettingsContent label[for='serverName']").removeAttr("style");
    	}
		if(exchangeSettingsNOK) {
			return;
		}
		$.ajax({
		    type: "POST",
		    url: "/portal/rest/exchange/settings",
		    data: JSON.stringify({
			        "serverName": $('.ExchangeEditSettingsContent #serverName').val(),
			        "domainName": $('.ExchangeEditSettingsContent #domainName').val(),
			        "username": $('.ExchangeEditSettingsContent #username').val(),
			        "password": $('.ExchangeEditSettingsContent #password').val()
			      }),
		    contentType: "application/json; charset=utf-8",
		    dataType: "json",
		    success: function(data){
				$(".ExchangeSettingsWindow").hide();
				$('.ExchangeSettingsButton').click();
			},
		    error: function(errMsg) {
		    	if(errMsg.statusText) {
			        alert(errMsg.statusText);
		    	} else {
		    		alert(errMsg);
		    	}
		    }
		});
	});
}
if ( document.addEventListener ) {
	window.addEventListener( "load", addExchangeButton, false );
} else if ( document.attachEvent ) {
	window.attachEvent( "onload", addExchangeButton );
}