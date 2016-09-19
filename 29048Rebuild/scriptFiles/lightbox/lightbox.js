/*
* Namespace the lightbox to make it more plugin like
*/
lightbox = (function()
{
	/*
	* Overlay is the translucent background beneath the box
	*/
	$overlay = $('<div id="overlay" class="hidden"></div>');
	/*
	* This is the outer wrapper for the lightbox content
	*/
	$lightbox = $('<div id="lightbox" class="hidden"></div>');
	/*
	* Content of lightbox is in its own container
	*/
	$content = $('<div id="lightboxContent"></div>');
	/*
	* Must have an icon to register a close action with
	*/
	$close = $('<a id="close">close</a>');
	/*
	* Add the contents and close to the lightbox wrapper
	*/
	$lightbox.append($content, $close);
	/*
	* Append lightbox and overlay to document
	*/
	$('body').append($overlay, $lightbox);
	/*
	* Set the click listener for the close event
	*/
	$close.on('click', function(e){
		e.preventDefault();
		method.close();
	});
	/*
	* Create an object that to hold the lightbox methods
	*/
	var method = {};
	/*
	* Method to initialize and open the lightbox
	*/
	method.open = function(content)
  	{
  		/*
  		* Add content that is passed in to the lightbox
  		*/
		$content.append(content);
		/*
		* Display the infobox and with the new content
		*/
		$lightbox.toggleClass("hidden");
		$overlay.toggleClass("hidden");
	};
	/*
	* Method to clean up and close the lightbox
	*/
	method.close = function ()
  	{
  		/*
  		* Hide the lightbox
  		*/
		$lightbox.toggleClass("hidden");
		$overlay.toggleClass("hidden");
		/*
		* Empty the content of the lightbox
		*/
		$content.empty();
		$(window).unbind('resize.modal');
	    /*
	    * This refreshes the browser to clear the globals and refresh memory
	    */
	    location.reload();
	};
	return method;
}());
