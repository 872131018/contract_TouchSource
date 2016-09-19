/*
*Begin logic for way finding app
*/
$(document).ready(function()
{
  /*
  * Only support a few global variables
  * Alpha is used to determine if x animation is complete
  * Omega is used to determine if y animation is complete
  * Index for iterating final segments during animation
  * Counter holds the current line segment to animate
  * Initialize the global around native canvas element
  * Coordinates used for positioning the infobox
  * company, building, suite are from database for infobox
  */
  alpha = false;
  omega = false;
  index = 1;
  counter = 0;
  canvas = new fabric.StaticCanvas("canvas");
  coordinates = {x: -1, y: -1};
  company = $("#mapvar_company").html();
  building = $("#mapvar_bldg").html();
  suite = $("#mapvar_suite").html();
  /*
  * Map empty building to ""
  */
  if(building == "empty")
  {
    building = "";
  }
  /*
  * Check for ' in the building name
  */
  if(building.indexOf("'") != -1)
  {
    building = building.replace("'", "&apos;")
  }
  /*
  *Load an xml with ajax that describes the path to draw
  *@TODO: also add error handling for failed requests or if the file is empty
  */
  $.get('paths.xml', function(xmlData, status)
  {
    /*
    *Define conditions that describe complete path
    *@TODO: need to write error handlers for when request fails or xml is malformed
    */
    if(status == 'success' && xmlData != '')
    {
      /*
      * Check if the building has a single suite labeled as all
      */
      if($(xmlData).find("building[name='"+building+"']").children("suite[name='all']").length == 1)
      {
        selector = "building[name='"+building+"'] > suite[name='all']";
      }
      else
      {
        selector = "building[name='"+building+"'] > suite[name='"+suite+"']";
      }
      /*
      * Valid nodes will have exactly 1 child node
      */
      if($(xmlData).find(selector).length != 1)
      {
        loadDefaultMap();
        return false;
      }
      /*
      * The xml gives the info box coordinates in a different format
      */
      var specialCaseArray = [$(xmlData).find(selector).attr("ibox")];
      /*
      * Move the infobox to its location and reveal
      */
      coordinates = createPoint(0, specialCaseArray);
      /*
      * Set the background image for the current path to draw
      */
      fabric.Image.fromURL($(xmlData).find(selector).attr("basemap"), function(oImg)
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
      * Select the building and suite from from the xml
      */
      $(xmlData).find(selector).each(function(suiteIndex, currentSuite)
      {
        /*
        * Get the color and width for the segments of the path
        */
        var color = '#' + $(xmlData).find("paths").attr("color");
        var strokeWidth = $(xmlData).find("paths").attr("width");
        strokeWidth = parseInt(strokeWidth);
        /*
        *Path to a suite is made up of an array of cartesian coordinates
        */
        var pointsArray = $(currentSuite).find("points").text().split(";");
        /*
        *Use the collection of segments to create the path
        */
        for(currentSegment in pointsArray)
        {
          /*
          *Stop 1 element short of the end of the array
          */
          if(parseInt(currentSegment) < pointsArray.length-1)
          {
            /*
            * Add a point at the beginning to round the end
            */
            canvas.add(
              createCircle(
                createPoint(currentSegment, pointsArray),
                color,
                strokeWidth
              ));
            /*
            *Create line from set of points and add it to canvass
            */
            canvas.add(
              createLine(
                createPoint(currentSegment, pointsArray),
                createPoint(currentSegment, pointsArray),
                color,
                strokeWidth
              ));
            /*
            *Add a point at the end to round the end
            */
            canvas.add(
              createCircle(
                createPoint(parseInt(currentSegment)+1, pointsArray),
                color,
                strokeWidth
              ));
          }
          /*
          * Put the finishing arrow on the final segment instead of circle
          */
          if(parseInt(currentSegment) == pointsArray.length-1)
          {
            canvas.add(
              createTriangle(
                createPoint(currentSegment, pointsArray),
                color
              ));
            /*
            *Convert final line circle pair to a rotation for the arrow
            */
            var pointA = createPoint(currentSegment, pointsArray);
            var pointB = createPoint(parseInt(currentSegment)-1, pointsArray);
            /*
            * Find the rotation angle based on the angle of the last line
            */
            var rotation =  Math.atan2(pointA.y - pointB.y, pointA.x - pointB.x) * (180 / Math.PI);
            /*
            *Rotate by 90 to account for the difference between the positive x axis and the triangle default rotation
            */
            rotation += 90;
            /*
            * Set rotation for the angle based on the computed value
            */
            triangle.angle = rotation;
          }
        }
      });
      /*
      * Begin animation of line segments in path
      */
      setTimeout(function()
      {
        animateSegment()
      }, 1000);
    }
    else
    {
      console.log("some kind of error I need to handle this!")
    }
  });
});
