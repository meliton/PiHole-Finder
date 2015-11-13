# PiHole-Finder
This is a software application to find the gateway and Pi-Hole Ad Blocker on your network.

Pi-Hole is a Raspberry Pi Ad Blocker for your network. <br />

Currently their project requires the user establish a static IP on the Pi.<br />

My twist on this concept is to allow a Pi-Hole to dynamically get an address, then use the PiHole-Finder software application to resolve the gateway and Pi-Hole IP. With this information, the user can then fire up their web browser and configure the Pi-Hole via a web interface.<br />

So, the workflow is:<br />
1. Install Raspian.<br />
2. Run the Pi-Hole Installation script.<br />
3. Run the PiHole Finder software application.<br /><br />

The Pi-Hole Finder depends on finding the Pi-Hole Ad Blocker's MAC address in the computer's arp cache. An easy way to do this is to have the Pi-Hole perform a ping sweep on the subnet it is on if it still has a dynamic address. I found <strong>fping</strong> to be lightest and easiest program to do this. Another possibility is to use the <strong>arping</strong> command. 
 






