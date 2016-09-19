// Wait until the DOM has loaded before querying the document
$(document).ready(function()
{
	/*
	* Restore the state of the window for better experience
	*/
	setTimeout(function()
	{
		$("html, body").scrollTop($.cookie("currentScroll"));
	}, 1000)
	/*
	* Each button image  will make ajax request for map
	*/
	$(".wayfinder").on('click', function(event)
	{
		/*
		* Save the window scroll state when the user finds the map
		*/
		$.cookie('currentScroll', window.scrollY);
		/*
		* Load the animated map and launch it in a modal
		*/
		$.get('animatedmap.asp?Companies='+$(this).data("company"), function(data)
		{
			/*
			* Save request url to create qr code from later
			*/
			map_url = window.location.origin+"/"+$(this).get(0).url;
			console.log(map_url)
			/*
			* Open lightbox with an object for settings
			*/
			lightbox.open(data);
		});
	});
});