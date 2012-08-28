/*
Snow Effect Script
Submitted by Altan d.o.o. (snow@altan.hr, http://www.altan.hr/snow/index.html)
Permission granted to Dynamicdrive.com to feature script in archive
For full source code to this script, visit http://dynamicdrive.com
Edited by Zolv.com for Virgin Holidays to fix various issues
*/

// Configure below to change number of snowflakes to render
var no = 10;

// Browser sniffer
var ns4up = (document.layers) ? 1 : 0;
var ie4up = (document.all) ? 1 : 0;
var ns6up = (document.getElementById&&!document.all) ? 1 : 0;

var dx, xp, yp;    // coordinate and position variables
var am, stx, sty;  // amplitude and step variables
var i, doc_width, doc_height;
var whichImg = 1, snowsrc;
var speed;	// 1 = smoothest, but more CPU; higher = more jerky but less CPU

dx = new Array(no);
xp = new Array(no);
yp = new Array(no);
am = new Array(no);
stx = new Array(no);
sty = new Array(no);
flake = new Array(no);

// Init the flakes
for (i = 0; i < no; ++i) {
	// 2 different types of flake
	whichImg = (whichImg==1?2:1);
	snowsrc="flake" + whichImg + ".gif";

	// Create the flakes in the page
	if (ns4up) {// set layers
		document.write("<layer name=\"dot"+ i +"\" left=\"15\" top=\"15\" visibility=\"show\"><img src='"+snowsrc+"' border=\"0\"></layer>");
	} else if (ie4up||ns6up) {
		document.write("<div id=\"dot"+ i +"\" style=\"position: absolute; z-index: "+ (i+100) +"; visibility: visible; top: 15px; left: 15px;\"><img src='"+snowsrc+"' border=\"0\"></div>");
	}
}

// Set internal variables for window width and height
function findWindowSize() {

	if (ns6up) {
		doc_width = window.innerWidth;
		doc_height = window.innerHeight;
	} else if (ie4up) {
		// Note that document.body is only available after page has loaded
		doc_width = document.body.clientWidth;
		doc_height = document.body.clientHeight;
	} else if (ns4up) {
		doc_width = self.innerWidth;
		doc_height = self.innerHeight;
	} else if (typeof(window.screen) == 'object') {	// Fallback, gets the screen width and height, not the window
		doc_width = window.screen.availWidth;
		doc_height = window.screen.availHeight;
	} else {
		// Some defaults, in case we can't work it out
		doc_width = 756;
		doc_height = 585;
	}
	//alert(doc_width + 'x' + doc_height);
}


// IE and NS6 main animation function
function snow() {
	for (i = 0; i < no; ++i) {
		yp[i] += sty[i];
		// Has this flake dropped off the bottom yet?
		if (yp[i] > doc_height - 50) {
			// Reset position
			xp[i] = Math.random()*(doc_width - am[i] - 30);
			yp[i] = 0;
			 // Set step variables for next fall
			stx[i] = Math.random()/10;
			sty[i] = (0.2 + Math.random()/3) * speed;
		}
		dx[i] += stx[i];
		// Reposition the flake
		if (ie4up) {
			flake[i].style.pixelTop = yp[i];
			flake[i].style.pixelLeft = xp[i] + am[i]*Math.sin(dx[i]);
		} else if (ns6up) {
			flake[i].style.top=yp[i] + "px";
			flake[i].style.left=xp[i] + am[i]*Math.sin(dx[i]) + "px";
		} else if (ns4up) {
			flake[i].top = yp[i];
			flake[i].left = xp[i] + am[i]*Math.sin(dx[i]);
		}
	}
	// Call ourselves again
	// Note we don't use setInterval on initialisation to allow ourselves to finish this call before starting the next (in case we're on a slow machine)
	setTimeout(snow, 10 * speed);
}

// Call once page has loaded - starts the snow going
function startSnowing() {

	findWindowSize();

	if (ie4up||ns4up) {
		speed = 1;	// Smooth
	} else if (ns6up) {
		speed = 5;	// Performance isn't great in Mozilla/NS with marquee scroller too - this makes it more jerky but uses less CPU
	}

	for (i = 0; i < no; ++i) {
		dx[i] = 0;
		// Initial random positioning of flakes
		xp[i] = Math.random()*(doc_width-50);
		yp[i] = Math.random()*doc_height;
		am[i] = Math.random()*10;	// Make flakes fall at different speeds
		stx[i] = Math.random()/10;	// set step variables
		sty[i] = (0.2 + Math.random()/3) * speed;
		// Various methods of accessing the flakes (factored out of loop in snow() to improve performance)
		if (ie4up) {
			flake[i] = document.all["dot"+i];
		} else if (ns6up) {
			flake[i] = document.getElementById("dot"+i);
		} else if (ns4up) {
			flake[i] = document.layers["dot"+i];
		}
	}

	// Go!
	if (ie4up||ns4up) {
		snow();
	} else if (ns6up) {
		setTimeout(snow, 1000);	// This wait is a workaround for Mozilla bug 165053
	}

}


// Set the events
//window.onload = startSnowing;	// For the vhols homepage, this is in the BODY onLoad instead
window.onresize = findWindowSize;	// In case the window is resized, we need to find its new size

