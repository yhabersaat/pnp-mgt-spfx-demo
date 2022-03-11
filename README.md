# pnp-mgt-spfx-demo

Demo SPFx solution presented at the Viva Connections and SharePoint Framework community call on March 10th, 2022.

1) Deploy mgt-spfx package on your tenant app catalog:
https://github.com/microsoftgraph/microsoft-graph-toolkit/releases
2) Clone this repository
3) In the command line run:
  - `npm install`
  - `gulp build`
  - `gulp bundle --ship`
  - `gulp package-solution --ship`
4) Add to your tenant app catalog and deploy
5) Go to SharePoint Admin Center and approve required API permissions