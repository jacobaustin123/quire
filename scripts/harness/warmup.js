// Launched once, invisibly, after install: macOS's first-launch security scan of a freshly built
// app takes several seconds, and this way it happens here instead of on the user's first open.
module.exports = async (ctx) => {
  await ctx.ready();
};
