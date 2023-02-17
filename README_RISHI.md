

## Development workflows

Your working master branch is `rishi-master`. 

To update `rishi-master` with `philc:master`:

1. Click sync "Sync fork" button on rkpatel33/sheetkeys
2. Open PR to merge `master` -> `rishi-master` (CONSIDER A REBASE WORKFLOW BELOW?)
3. Look at changes in the PR diff.

How to keep your Git-Fork up to date
https://stefanbauer.me/articles/how-to-keep-your-git-fork-up-to-date



## Development

Google docs on debugging contetnt script:

https://developer.chrome.com/docs/extensions/mv3/tut_debugging/#debug_cs


```javascript

var menuItemNodes = document.querySelectorAll(".menu-button");
var menuItems = Array.from(menuItemNodes);
menuItems = menuItems.filter(x => x.innerText.includes('Albert'));

if (menuItems.length > 0) {
  var customMenu = menuItems[0];
  window.KewyboardUtils.simulateClick(customMenu);
}

  const el = Array.from(document.querySelectorAll(".menu-button"))
    .find(el => el.textContent.includes('Albert'));


SheetActions.clickMenuCustomItem('Albert', '$0.00');

SheetActions.createCustomMenus()

```
