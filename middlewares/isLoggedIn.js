function isLoggedIn(req, res, next) {
  if (req.isAuthenticated && req.isAuthenticated()) {
    return next();
  }
  
  const isAjax = req.xhr || req.headers.accept?.includes("application/json") || req.headers["x-requested-with"] === "XMLHttpRequest";
  if (isAjax) {
    return res.status(401).json({ status: false, message: "Session expired. Please log in again.", redirect: "/login" });
  }

  res.redirect("/login");
}

module.exports = { isLoggedIn };
