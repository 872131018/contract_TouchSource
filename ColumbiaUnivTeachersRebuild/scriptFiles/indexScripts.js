function loadDefaultMap()
{
  /*
  * Set the background image for the current path to draw
  */
  fabric.Image.fromURL("defaultmap.jpg", function(oImg)
  {
    /*
    * Now that the image is loaded its size can be found for setting the canvas
    */
    canvas.setWidth(oImg.width);
    canvas.setHeight(oImg.height);
    canvas.renderAll();
    /*
    * Image has to be added as callback because of lazy loading
    * While the image is loaded, other elements are created and added
    * By the time image is added other things are under it
    * In order for the image to render correctly it must be the bottommost item
    */
    canvas.add(oImg);
    canvas.moveTo(oImg, 0);
  });
  /*
  * Set up infobox and display it
  */
  infobox.open(
  {
    header: "There is no information for that destination.",
    status: "error",
    left: "400px",
    top: "200px"
  });
  /*
  * Also animate the reveal of the infobox
  */
  $("#infobox-wrapper").animate(
    {
      opacity: 0.8
    }, 1500);
}
/*
* Begin the logic for each segment of the path
*/
function animateSegment()
{
  /*
  * Canvas objects consist of circle, line, circle groups in each segment
  */
  if(index < canvas.getObjects().length-2)
  {
    /*
    *Compute delay for next line segment from distance between circles
    */
    var sideA = Math.abs(canvas.item(index).left - canvas.item(index+2).left);
    var sideB = Math.abs(canvas.item(index).top - canvas.item(index+2).top);
    var length = Math.sqrt((sideA*sideA) + (sideB*sideB));
    /*
    * Delay is shifted based on a set rate of animation (100px/second)
    */
    var rate = 100;
    var delta = length/rate;
    var delay = 1000*delta;
    /*
    * Show the circle from which the line will start
    */
    canvas.item(index).animate("opacity", 1.0,
    {
      duration: 1,
      onChange: canvas.renderAll.bind(canvas)
    });
    /*
    * Reveal the line 
    */
    canvas.item(index+1).animate("opacity", 1.0,
    {
      duration: 1,
      onChange: canvas.renderAll.bind(canvas)
    });
    /*
    * Animate the line to the second circle
    */
    canvas.item(index+1).animate("x2", canvas.item(index+2).left,
    {
      duration: delay,
      onChange: canvas.renderAll.bind(canvas),
      onComplete: function()
      {
        /*
        * Reveal the finishing circle when complete 
        */
        alpha = !alpha;
        if(alpha && omega)
        {
          canvas.item(counter).animate("opacity", 1.0,
          {
            duration: 1,
            onChange: canvas.renderAll.bind(canvas)
          });
        }
      }
    });
    canvas.item(index+1).animate("y2", canvas.item(index+2).top,
    {
      duration: delay,
      onChange: canvas.renderAll.bind(canvas),
      onComplete: function()
      {
        /*
        * Reveal the finishing circle when complete 
        */
        omega = !omega;
        if(alpha && omega)
        {
          canvas.item(counter).animate("opacity", 1.0,
          {
            duration: 1,
            onChange: canvas.renderAll.bind(canvas)
          });
        }
      }
    });
    /*
    * Continue animation in a daisy chain fashion
    */
    setTimeout(function()
    {
      animateSegment()
    }, delay);
    /*
    * Increment by three because a segment is a group of circle, line, circle objects
    */
    index += 3;
  }
  else if(index == canvas.getObjects().length-1)
  {
    /*
    * On final segment reveal finish arrow
    */
    canvas.item(index).animate("opacity", 1.0,
    {
      duration: delay/2,
      onChange: canvas.renderAll.bind(canvas)
    });
    /*
    * If coordinates are set create and display the infobox
    */
    if(!isNaN(coordinates.x) && !isNaN(coordinates.y))
    {
      /*
      * Open the infobox with some settings
      */
      infobox.open(
      {
        header: company+" <br> "+suite,
        status: "success",
        left: coordinates.x+"px",
        top: coordinates.y+"px"
      });
      /*
      * Reveal infobox
      */
      $("#infobox-wrapper").animate(
      {
        opacity: 1.0
      }, 2000);
    }
    /*
    * Init the tracer function by resetting the index and delay
    */
    index = 1;
    pathTrace();
  }
}
function pathTrace()
{
  /*
  * Canvas objects consist of line and circle pairs, stop 1 short of end
  */
  if(index < canvas.getObjects().length-2)
  {
    /*
    *Compute delay for next line segment
    */
    var sideA = Math.abs(canvas.item(index).left - canvas.item(index+2).left);
    var sideB = Math.abs(canvas.item(index).top - canvas.item(index+2).top);
    var length = Math.sqrt((sideA*sideA) + (sideB*sideB));
    /*
    * Delay is shifted based on a set rate of animation (100px/second)
    */
    var rate = 100;
    var delta = length/rate;
    var delay = 1000*delta;
    /*
    * Make a local copy of the current line
    */
    var currentLine = canvas.item(index+1);
    /*
    * Set the color and size of the tracer point
    */
    var color = "#ff0000";
    var strokeWidth = "14";
    /*
    * Extract data from current line and transform it for createpoint()
    */
    var specialCaseArray = [[currentLine.x1, currentLine.y1].join()];
    /*
    * Create the circle and add it to the canvas
    */
    canvas.add(
      createCircle(
        createPoint(0, specialCaseArray),
        color,
        strokeWidth
      )
    );
    /*
    * Tracer pushed onto top of stack
    */
    var tracer = canvas.item(canvas.getObjects().length-1);
    /*
    * Animate the point from one end of the line to the other
    */
    tracer.animate("left", currentLine.x2,
    {
      duration: delay,
      onChange: canvas.renderAll.bind(canvas)
    });
    tracer.animate("top", currentLine.y2,
    {
      duration: delay,
      onChange: canvas.renderAll.bind(canvas)
    });
    tracer.animate("opacity", 1.0,
    {
      duration: 1,
      onChange: canvas.renderAll.bind(canvas)
    });
    setTimeout(function()
    {
      /*
      * Pop tracer off canvas stack
      */
      var tracer = canvas.item(canvas.getObjects().length-1);
      canvas.remove(tracer);
      pathTrace();
    }, delay);
    /*
    * Increment by two because the lines are paired with points
    */
    index += 3;
  }
  else if(index == canvas.getObjects().length-1)
  {
    /*
    * Reset index to restart animation
    */
    index = 1;
    /*
    * Animation will continue forever, just needs to reset
    */
    setTimeout(function()
    {
      pathTrace();
    }, delay);
  }
}
function createLine(passedPointA, passedPointB, passedColor, passedWidth)
{
  return new fabric.Line(
    [passedPointA.x, passedPointA.y, passedPointB.x, passedPointB.y],
    {
      originX: "center",
      originY: "center",
      //opacity: 0,
      opacity: 1.0,
      strokeWidth: parseInt(passedWidth),
      stroke: passedColor
    }
  );
}
function createTriangle(passedPoint, passedColor)
{
  return triangle = new fabric.Triangle(
  {
    width: 16,
    height: 16,
    originX: "center",
    originY: "center",
    fill: passedColor,
    opacity: 0,
    left: parseInt(passedPoint.x),
    top: parseInt(passedPoint.y)
  });
}
function createCircle(passedPoint, passedColor, passedWidth)
{
  return new fabric.Circle(
  {
    radius: parseInt(passedWidth) / 2,
    originX: "center",
    originY: "center",
    fill: passedColor,
    opacity: 0,
    left: parseInt(passedPoint.x),
    top: parseInt(passedPoint.y)
  });
}
function createPoint(passedIndex, passedArray)
{
  var point = {};
  point.x = parseInt(passedArray[passedIndex].split(",")[0]);
  point.y = parseInt(passedArray[passedIndex].split(",")[1]);
  return point;
}
