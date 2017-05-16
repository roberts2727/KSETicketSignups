## kse-ticket-signups

This web part will connect a games list with an attendees list to accomplish a ticket signup system that allows for an alloted number of tickets per game to be assigned and drawn from. Once the remaining seats reaches 0 the game will close. You can close a game at any time by changing the remaining seats to 0, or extend allotment by adding to the alloted field. The Games list Must be titled "Games", and the attendees list must be titled "Attendees". The web part will prompt the user for the number of tickets the would like to register for, their Perferred Flash Seats account, and if they have any special requests such as Handicap Accessable Seats.

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

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
