<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8"/>
    <title>THEOPlayer - Ramp Utilization PoC</title>
    <script type="text/javascript" src="https://cdn.myth.theoplayer.com/0906c099-1722-4fbb-88dd-87bfa3148d98/THEOplayer.js"></script>
    <link rel="stylesheet" type="text/css" href="https://cdn.myth.theoplayer.com/0906c099-1722-4fbb-88dd-87bfa3148d98/ui.css"/>
    <script type="text/javascript" src="//ramp-poc/rampapi.js"></script>
</head>
<body>
	<p id="status">loading...</p>
	<div class="theoplayer-container video-js theoplayer-skin vjs-16-9"></div>
	<script>
		const params = {
			i: {
				onDone: done,
				allowFallback: true,
				fallbackUrl: false
			},
			r: {
				host: 'ramp-demo.multicast-receiver-altitudecdn.net',
				maddr: '.....',
				verbose: false,
				monitorStatus: true,
				pingTimeMs: 5000			
			},
			o: {}
		};
		const receiver = new RcvrInterface(params.i, params.r, params.o);

		var player = null;
		var ignorePlayerWaitingEvent = false;
		
		function restartPlayer() {
			if (player != null) {
				ignorePlayerWaitingEvent = true;
				player.source = player.source;
				player.play();
				console.log(formatAMPM(Date.now()) + ":: THEOplayer re-connected + started to play");
			}
		}
		
		function done(ifc, obj) {
			console.log(formatAMPM(Date.now()) + ":: RcvrInterface 'done'");
			console.log(obj);
			var rewarmOccurred = false;
			if (obj.url === false) {
				document.getElementById('status').innerText = 'using fallback';
				obj.url = originalHLSstreamAddressLink;
			}
			else {
				document.getElementById('status').innerText = 'using ramp';
			}

			let element = document.querySelector('.theoplayer-container');
			if (player == null) {
				player = new THEOplayer.Player(element, {
					libraryLocation: 'https://cdn.myth.theoplayer.com/0906c099-1722-4fbb-88dd-87bfa3148d98'
				});
				player.addEventListener('error', function(errorEvent) {
					console.log(formatAMPM(Date.now()) + ":: THEOPlayer - error:: " + errorEvent);
				});
				
				player.addEventListener('waiting', playerWaiting);
			}

			player.source = {
				sources: [
					{
						'src': obj.url,
						'type': 'application/x-mpegurl'
					}
				]
			};
			player.preload = 'auto';
		}
		
		function playerWaiting() {
			//regular waiting do nothing
			console.log(formatAMPM(Date.now()) + ":: THEOPlayer - waiting event");
			if (!player.seeking && !player.paused) {
				if (!ignorePlayerWaitingEvent) {
					console.log(formatAMPM(Date.now()) + ":: THEOPlayer - stalled");									
					restartPlayer();
				} else {
					ignorePlayerWaitingEvent = false;
				}
			}
		}
		
		function formatAMPM(milliseconds) {
			var date = new Date(milliseconds);
			return date.toDateString() + " - " + date.toLocaleTimeString();	
		}
	</script>
	</body>
</html>