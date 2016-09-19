/*
* Namespace the infobox to make it more plugin like
*/
infobox = (function()
{
	/*
	* This is the outer wrapper for the lightbox content
	*/
	$container = $('<div class="infobox-container"></div>');
	/*
	* Triangles are used for visual effect on banner
	*/
	$triangleLeft = $('<div class="triangle left"></div>');
	$triangleRight = $('<div class="triangle right"></div>');
	/*
	* Create a container for the icon
	*/
	$iconContainer = $('<div id="iconContainer"></div>')
	/*
	* Icon to access the qr code section
	*/
	if(icon_tone == "light")
	{
		$qrIcon = $('<img id="qrIcon" src="scriptFiles/infobox/QR-Scan-Light.png">');
	}
	else
	{
		$qrIcon = $('<img id="qrIcon" src="scriptFiles/infobox/QR-Scan-Dark.png">');
	}
	/*
	* This is the actual infobox component
	*/
	$infobox = $('<div class="infobox"></div>');
	/*
	* Wrap the header in an h3 tag
	*/
	$headerWrapper = $('<h3></h3>');
	/*
	* This is the banner of the infobox wrapped in h3 tags
	*/
	$header = $('<span id="infoboxHeader"></span>');
	/*
	* This is a div to hold the content beneath the banner
	*/
	$content = $('<div class="hidden" id="infoboxContent"></div>');
	/*
	* Container to hold the QR code
	*/
	$qrContainer = $('<div id="qrContainer"></div>');
	/*
	* This is where the QR code will go
	*/
	$qrCode = $('<div id="qrcode"></div>');
	/*
	* Build structure of infobox based on pieces
	*/
	$container.append($triangleLeft);
	$container.append($triangleRight);
	$iconContainer.append($qrIcon)
	$container.append($iconContainer);
	$infobox.append($headerWrapper);
	$qrContainer.append($qrCode);
	$content.append($qrContainer);
	$infobox.append($content);
	$container.append($infobox);
	$headerWrapper.append($header);
	/*
	* Append infoxbox to wrapper in document
	*/
	$('#infobox-wrapper').append($container);
	/*
	* Set the click listener for the expanding switch
	*/
	$qrIcon.on('click', function(event)
	{
		event.preventDefault();
		/*
		* Expand code goes here if there is any
		*/
		$qrIcon.toggleClass("flip");
		$content.toggleClass("hidden");
	});
	/*
	* Create an object that to hold the infobox methods
	8*/
	var method = {};
	/*
	* Method to initialize and open the lightbox
	*/
	method.open = function(settings)
  	{
  		/*
      * Set the content of the header
      */
      $header.html(settings.header);
      /*
      * Target the banner by setting the h3 it contains
      */
      $header.parent().addClass(settings.status);
      /*
      * Move the infobox my targeting the parent wrapper
      */
      $container.parent().css("left", settings.left);
      $container.parent().css("top", settings.top);
      /*
	  * Create a qr code object in the #qrcode element
	  */
	  var qrcode = new QRCode("qrcode",
	  {
	    width: 180,
	    height: 180,
	    colorDark : "#005baa",
	    colorLight : "#ffffff",
	    correctLevel: QRCode.CorrectLevel.L
		});
	  /*
	  * Generate a new code from a url
	  */
	  qrcode.makeCode(map_url);
	};
	/*
	* Could close the infobox but not really a reason to
	*/
	method.close = function ()
  	{

	};
	/*
	* Allows the infobox to call methods like a class
	*/
	return method;
}());
