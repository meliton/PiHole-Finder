# PiHole-Finder
This is a software application to find the gateway and Pi-Hole Ad Blocker on your network.

Pi-Hole is a Raspberry Pi Ad Blocker for your network.

Currently their project requires the user establish a static IP on the Pi.

My twist on this concept is to allow a Pi-Hole to dynamically get an address, then use the PiHole-Finder software application to resolve the gateway and Pi-Hole IP. With this information, the user can then fire up their web browser and configure the Pi-Hole via a web interface.

So, the workflow is:
1. Install Raspian.<br />
2. Run the Pi-Hole Installation script.
3. Run the PiHole software application.


