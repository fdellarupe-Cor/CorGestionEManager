/*  ContentFlowAddOn_circle, version 1.0
 *  (c) 2008 - 2010 Sebastian Kutsch
 *  <http://www.jacksasylum.eu/ContentFlow/>
 *
 *  This file is distributed under the terms of the MIT license.
 *  (see http://www.jacksasylum.eu/ContentFlow/LICENSE)
 */
new ContentFlowAddOn ('circle', {


    conf: {
        showArrows: true,
        scrollSpeed: 0.1
    },

    init: function() {
    },

    onloadInit: function (flow) {
    },

    afterContentFlowInit: function (flow) {
        var conf = flow.getAddOnConf('circle');
        if (conf.showArrows) {

            var H = flow.Flow.clientHeight;
            var W = flow.Flow.clientWidth;
            var o = 0.3;
            var t = H*0.425+"px";
            var h = H*0.15+"px";;
            var m = W*0.32+"px";
            var imgPath = ContentFlowGlobal.BaseDir+"img";

            var al = new Image();
            al.src = imgPath+"/arrow_left.png";
            al.style.height = h;
            al.style.position = "absolute";
            al.style.left = m;
            al.style.top = t;
            al.style.opacity = o;
            flow.Container.appendChild(al);
            flow.t1 = null;
            al.onmouseover = function () {
                flow.t1 = window.setInterval(function () {flow.moveToPosition(flow._targetPosition-conf.scrollSpeed)}, 100);
                al.style.opacity = 1;
            }
            al.onclick = function () {
                flow.moveToPosition(--flow._targetPosition)
            }
            al.onmouseout = function () {
                window.clearInterval(flow.t1);
                flow.moveToPosition(Math.floor(flow._targetPosition));
                al.style.opacity = o;
            }

            var ar = new Image();
            ar.src = imgPath+"/arrow_right.png";
            ar.style.height = h;
            ar.style.position = "absolute";
            ar.style.right = m;
            ar.style.top = t;
            ar.style.opacity = o;
            flow.Container.appendChild(ar);
            flow.t2 = null;
            ar.onmouseover = function () {
                flow.t2 = window.setInterval(function () {flow.moveToPosition(flow._targetPosition+conf.scrollSpeed)}, 100);
                ar.style.opacity = 1;
            }
            ar.onclick = function () {
                flow.moveToPosition(++flow._targetPosition)
            }
            ar.onmouseout = function () {
                window.clearInterval(flow.t2);
                flow.moveToPosition(Math.ceil(flow._targetPosition));
                ar.style.opacity = o;
            }
        }
    },

	ContentFlowConf: {
        scaleFactorLandscape: "max",      // scale factor of landscape images ('max' := height= maxItemHeight)
        fixItemSize: true,
        relativeItemPosition: "center", // align top/above, bottom/below, left, right, center of position coordinate
        visibleItems: 8,               // how man item are visible on each side (-1 := auto)
        endOpacity: 0.5,                  // opacity of last visible item on both sides

        calcSize: function (item) {
            var rP = item.relativePosition;
            var h = Math.abs(rP) <= 0.5 ? 0.5 - Math.abs(rP)/2 : 0.25;
            var w = h;
            return {width: w, height: h};
        },

        calcCoordinates: function (item) {
            var rP = item.relativePosition;
            var rPN = item.relativePositionNormed;
            var vI = this.conf.visibleItems;
            var f = (rP/vI+0.5)*Math.PI;
            var x = Math.cos(f)*0.5;
            var y = Math.sin(f)+2*(1-Math.abs(rPN))+0.25 - (Math.abs(rP) <= 0.5 ? 2.5* (0.5-Math.abs(rP)) : 0) ;
            return {x: x, y: y};
        },

        calcZIndex: function (item) {
            return -Math.abs(item.relativePositionNormed);
        },

        calcFontSize: function (item) {
            return item.size;
        },

        calcOpacity: function (item) {
            return 1 - Math.abs(item.relativePositionNormed/2);
        }

    }

});
