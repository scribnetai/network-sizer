# Network Sizer

TOR & FC SAN switch planner for VMware refreshes. Fourth app in the presales SE sizer suite
([server-sizer](https://server-sizer.scribnet.io/),
[storage-sizer](https://storage-sizer.scribnet.io/),
[rvtools-analyzer](https://rvtools-analyzer.scribnet.io/)).

**Live:** https://network-sizer.scribnet.io/

## What it does

Sizes top-of-rack Ethernet switches and Fibre Channel SAN switches from the same
demand signal the other sizers use — your host count.

1. **Load demand** — drop an RVTools `.xlsx` (reads the `vHost` tab), enter host
   groups manually, or run the 21-host demo. Optionally list existing switches
   to get a net-new vs reuse view.
2. **Configure fabrics** — per-host port profile (mgmt / data-vMotion /
   storage-IP ports × speeds, FC HBAs × speed), TOR preset (Cisco Nexus,
   Arista, Dell, 100G shared-pool, or custom), uplinks per switch,
   oversubscription target, A/B dual-homing, N+1 spare, 4×25G breakout toggle.
   FC side: Brocade G720/G730 or Cisco MDS 9132T/9148T/9396T (or custom), ISL
   reserve, array count × target ports, dual Fabric A/B.
3. **Network plan** — BOM, port utilization, oversubscription vs target,
   per-fabric worked math, SE findings (talking points), and a downloadable
   standalone HTML report.

## Sizing math

- Ethernet: downlink ports = hosts × Σ ports/host; switches =
  `ceil(ports ÷ usable downlinks/switch)`, rounded to pairs for A/B dual-homing,
  plus optional N+1 spare. Oversubscription = total downlink bandwidth ÷ total
  uplink bandwidth.
- 100G shared-pool preset: uplinks carve out of the 64-port pool; with 4×25G
  breakout, ≤25G host connections ride four-per-port.
- FC: device ports = host HBAs + array targets, split across Fabric A/B;
  usable ports/switch = total − ISL reserve; switches per fabric =
  `ceil(device ports ÷ usable)`.
- Ports needing faster downlinks than the TOR supports are flagged as unserved
  rather than silently sized.

## Privacy

100% client-side. The spreadsheet is parsed in your browser with a vendored
copy of SheetJS (works offline) — nothing is uploaded, stored, or sent anywhere.

## Tests

Pure sizing functions are exported for node:

```bash
node /tmp/net-tests.js   # 32 assertions over sizeEthernet / sizeFC / net-new
```
