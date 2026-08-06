function isCompAdmin(req, res, next) {
  const isAjax = req.xhr || req.headers.accept?.includes("application/json") || req.headers["x-requested-with"] === "XMLHttpRequest";

  if (req.isAuthenticated && req.isAuthenticated()) {
    const user = req.user;
    const roleNum = Number(user?.role);
    if (user && (roleNum === 2 || roleNum === 1)) {
      return next();
    } else {
      if (isAjax) {
        return res.status(403).json({ status: false, message: "Unauthorized access.", redirect: "/unauthorized" });
      }
      return res.redirect("/unauthorized");
    }
  } else {
    if (isAjax) {
      return res.status(401).json({ status: false, message: "Session expired. Please log in again.", redirect: "/login" });
    }
    return res.redirect("/login");
  }
}

module.exports = { isCompAdmin };
