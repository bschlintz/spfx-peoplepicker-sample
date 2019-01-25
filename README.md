## spfx-peoplepicker-sample

Sample SPFx webpart containing a modified version of the [PnP SPFx React People Picker](https://sharepoint.github.io/sp-dev-fx-controls-react/controls/PeoplePicker/) control. The People Picker control has been modified to showcase 4 different implementation options to filter users out from the auto-complete suggestions list.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

---

Microsoft provides programming examples for illustration only, without warranty either expressed or
implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
for a particular purpose. We grant You a nonexclusive, royalty-free right to use and modify the 
Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that
You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the
Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which
the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers
from and against any claims or lawsuits, including attorneys' fees, that arise or result from the
use or distribution of the Sample Code.

This sample assumes that you are familiar with the programming language being demonstrated and the 
tools used to create and debug procedures. Microsoft support professionals can help explain the 
functionality of a particular procedure, but they will not modify these examples to provide added 
functionality or construct procedures to meet your specific needs. if you have limited programming 
experience, you may want to contact a Microsoft Certified Partner or the Microsoft fee-based consulting 
line at (800) 936-5200. 
