var $ = function(e) {
	return document.querySelector(e);
};

var frequency = function() {
	osc.frequency.value = this.value;
};

var volume = function() {
	gainNode.gain.value = this.value;
};

$('.start').onclick = function() {
	analyser.connect(audio.destination);
	osc.start();
};

$('.stop').onclick = function() {
	analyser.disconnect(audio.destination);
};

$('.frequency').oninput = frequency;
$('.volume').oninput = volume;
$('.waves').onclick = function(e) {
	osc.type = e.target.id;
};

/////////////////////////////////////

var canvas = $('canvas');
var ctx = canvas.getContext('2d');

var audio = new window.AudioContext();
var osc = audio.createOscillator();
var gainNode = audio.createGain();
var analyser = audio.createAnalyser();

osc.frequency.value = 131;
osc.connect(gainNode);
gainNode.connect(analyser);

///

analyser.fftSize = 1024;
var bufferLength = analyser.fftSize;
var dataArray = new Uint8Array(bufferLength);
analyser.getByteTimeDomainData(dataArray);

function draw() {

      drawVisual = requestAnimationFrame(draw);
      analyser.getByteTimeDomainData(dataArray);

      ctx.fillStyle = '#111';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.lineWidth = 2;
      ctx.strokeStyle = 'rgba(207,63,63,.9)';
      ctx.beginPath();

      var sliceWidth = canvas.width * 1.0 / bufferLength;
      var x = 0;

      for(var i = 0; i < bufferLength; i++) {
   
        var v = dataArray[i] / 128.0;
        var y = v * canvas.height/2;

        if(i === 0) {
          ctx.moveTo(x, y);
        } else {
          ctx.lineTo(x, y);
        }

        x += sliceWidth;
      }

      ctx.lineTo(canvas.width, canvas.height/2);
      ctx.stroke();
};

draw();