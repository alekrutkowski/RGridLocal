\name{openRGrid}
\alias{openRGrid}
\title{Open RGrid-local}
\usage{
openRGrid(host = "127.0.0.1", port = 0L,
          launch.browser = interactive(), block = interactive(), quiet = FALSE)
}
\arguments{
\item{host}{Interface to bind. The default only accepts connections from the local machine.}
\item{port}{TCP port to try first. Use \code{0} or \code{NULL} to choose an available random port. If a positive port is unavailable, the next 50 ports are tried in order.}
\item{launch.browser}{Logical or function. When \code{TRUE}, open the browser after the server starts. When a function, it is called with the URL.}
\item{block}{Keep servicing requests until the user interrupts R. Set to \code{FALSE} to return a server handle immediately.}
\item{quiet}{Suppress startup messages.}
}
\description{
Start a localhost RGrid session whose spreadsheet formulas are evaluated by the standard R interpreter that is running this package.
}
\value{
Invisibly returns a server handle when \code{block = FALSE}; otherwise the server runs until interrupted.
}
\examples{
\dontrun{
openRGrid()
}
}
