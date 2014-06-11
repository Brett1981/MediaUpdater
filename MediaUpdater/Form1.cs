//Sam Hissaund 2014
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Net;
using System.Web;
using System.Web.Script.Serialization;
using System.Collections;

/*Does something like:
//Read all dirs

//For each dir:
//Check for a file with the imdb movie ID as the filename Done
//If available, ignore. Done
//If no file, do a search on IMDB and MovieDB for data.
//Array the data.
//Add file to dir with the imdb movie ID as the filename


//Connect to database on home.cplsyx.com
//Log in
//Update the DB with the list of new movies (in the array)
//Close DB
 * */

namespace MediaUpdater
{
    public partial class Form1 : Form
    {
        private string moviePath = "";
        private string tvPath = "";
        private string imdbAPI = "";
        private string tvdbAPI = "";
        private string moviedbAPI = "";
        private string moviedbAPIkey = "";
        private string localdbName = "";
        private string localdbUser = "";
        private string localdbPW = "";
        private string localdbHost = "";

        /*
         * Update the log message window on the app. This is thread safe so we can run the work in another thread (so the form isn't blocked)
         */
        private void log(String str, bool newLine = true)
        {
            if (output.InvokeRequired)
            {
                MethodInvoker invoker = () => log(str, newLine);
                Invoke(invoker);
            }
            else
            {
                if (newLine)
                {
                    output.AppendText(str + "\r\n");
                }
                else
                {
                    output.AppendText(str);
                }
            }
            
            
        }

        /* 
         * Load the saved arguments so we can reuse this application again without needing to recompile if the directories/API URLs change
         */
        private void loadArgs() 
        {
            this.log(">Loading arguments...");
            this.moviePath = Properties.Settings.Default.moviePath;
            this.tvPath = Properties.Settings.Default.tvPath;
            this.imdbAPI = Properties.Settings.Default.imdbAPI;
            this.tvdbAPI = Properties.Settings.Default.tvdbAPI;
            this.moviedbAPI = Properties.Settings.Default.moviedbAPI;
            this.moviedbAPIkey = Properties.Settings.Default.moviedbAPIkey;
            this.localdbName = Properties.Settings.Default.localdbName;
            this.localdbUser = Properties.Settings.Default.localdbUser;
            this.localdbPW = Properties.Settings.Default.localdbPW;
            this.localdbHost = Properties.Settings.Default.localdbHost;
            this.log("moviePath: " + moviePath);
            this.log("tvPath: " + tvPath);
            this.log("imdbAPI: " + imdbAPI);
            this.log("tvdbAPI: " + tvdbAPI);
            this.log("moviedbAPI: " + moviedbAPI);
            this.log("moviedbAPIkey: " + moviedbAPIkey);
            this.log("localdbName: " + localdbName);
            this.log("localdbUser: " + localdbUser);
            this.log("localdbPW: " + localdbPW);
            this.log("localdbHost: " + localdbHost);
            this.log(">Arguments loaded.");
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.runMinimized)
            {
                this.WindowState = FormWindowState.Minimized;
                notifyIcon1.ShowBalloonTip(3);
                this.ShowInTaskbar = false;
            }
            this.log(">Initialising...");
            this.loadArgs();
            this.log(">Initialising complete.");
            Thread t = new Thread(processController); //Run the work in a new thread to stop the form getting blocked
            t.Start();
        }
        
        /*
         * Overseeing process that I am using to call others, mainly for neatness and debugging
         */
        private void processController()
        {
            this.log(">Processing Movies");
            this.processMovies();
            this.log(">Processing Movies complete");
        }

        /*
         * Process movies into DB
         */
        private void processMovies()
        {
            string[] folders = System.IO.Directory.GetDirectories(@moviePath);
            string[] newMovies = new string[folders.Length]; //Get the directories and create an array ready for new movies, at least the size of the DIRs in the movies folder (in case this is a fresh run)
            int newMovieCount = 0;
            foreach (string folder in folders) //Iterate over all the movies in the folder
            {
                try
                {

                    string[] fileList = System.IO.Directory.GetFiles(folder); //Checking for the IMDB identifier file we will be using
                    bool movieExists = false;
                    foreach (string pathName in fileList)
                    {
                        string fileName = Path.GetFileName(pathName);
                        if (fileName.StartsWith("imdb")) //If changing this ident, also ensure the "createImdbIdentFile" function is updated
                        {                            
                            movieExists = true;
                            break;
                        } 
                       
                    }
                    if (!movieExists) //if file not found then movie must be new and we'll add it to the array for processing
                    {
                        newMovies[newMovieCount] = Path.GetFileName(folder);
                        newMovieCount++;
                        this.log(folder + " is a new movie");
                    }
                }
                catch (Exception e)
                {
                    this.log(@moviePath + folder + " generated exception: " + e.Message);
                }
            }
            this.log("New movies detected: "+newMovieCount);
            if (newMovieCount > 0)
            {
                foreach (string movieTitle in newMovies) //Got our new movies, time to process them.
                {
                    Dictionary<string, string> SQLQueryVals = checkAPI(movieTitle, "movie"); //Grab the details we need for the DB
                    if (SQLQueryVals == null || SQLQueryVals.Count < 9)
                        continue;

                    string query = "INSERT INTO Movies ([imdbID], [imdbTitle], [imdbYear], [imdbGenre], [imdbRating], " +
                    "[imdbRuntime], [imdbPlot], [moviedbImage], [moviedbID]) VALUES (@iID, @iTitle, @iYear, @iGenre, @iRating, @iRunTime, @iPlot, @mImage, @mID);";


                    SqlCommand sqlquery = new SqlCommand(query); //Insert the query params into the query. This protects against injection attack.
                    sqlquery.Parameters.AddWithValue("@iID", SQLQueryVals["imdbID"]);
                    sqlquery.Parameters.AddWithValue("@iTitle", SQLQueryVals["imdbTitle"]);
                    sqlquery.Parameters.AddWithValue("@iYear", SQLQueryVals["imdbYear"]);
                    sqlquery.Parameters.AddWithValue("@iGenre", SQLQueryVals["imdbGenre"]);
                    sqlquery.Parameters.AddWithValue("@iRating", SQLQueryVals["imdbRating"]);
                    sqlquery.Parameters.AddWithValue("@iRunTime", SQLQueryVals["imdbRuntime"]);
                    sqlquery.Parameters.AddWithValue("@iPlot", SQLQueryVals["imdbPlot"]);
                    sqlquery.Parameters.AddWithValue("@mImage", SQLQueryVals["moviedbImage"]);
                    sqlquery.Parameters.AddWithValue("@mID", SQLQueryVals["moviedbID"]);

                    if (SQLInsert(sqlquery))
                    {
                        createImdbIdentFile(movieTitle, SQLQueryVals["imdbID"]);
                    }


                }
            }
            return;

        }

        private void createImdbIdentFile(string folder, string imdbID)
        {
            log("Creating IdentFile for " + folder + "... ", false);
            try
            {
                //Create the file, make sure it starts with "imdb" to identify it
                string identFile = @moviePath + "\\" + folder + "\\imdb" + imdbID;
                //File.Create(identFile);
                log("Done.");
            }
            catch (Exception e)
            {
                log("Error.");
                log(e.Message);
            }

        }

        private bool SQLInsert(SqlCommand command)
        {
            log("Inserting into database... ", false);
            try
            {



                using (SqlConnection connection = new SqlConnection("User ID="+localdbUser+";" +
                                           "Password="+localdbPW+";"+
                                           "Server=tcp:" + localdbHost + ",1433;" +
                                           "Trusted_Connection=false;" + //This has to be false, otherwise connection will attempt to use Windows account credentials.
                                           "Database="+localdbName+"; " +
                                           "connection timeout=30"))
                {
                    
                    command.Connection = connection;
                    connection.Open();
                    command.ExecuteNonQuery();                    
                    log("Done.");
                    return true;
                }
            }
            catch (Exception e)
            {
                log("Failed!");
                log("Error with SQL insert:" + e.Message);
                return false;
            }
           
        }

        /*
         * Takes the response from the GET and processes these into the values we need for our SQL query. 
         */
        public Dictionary<string, string> checkAPI(string title, string type)
        {
            if (type == "movie")
            {

                var SQLQueryValues = new Dictionary<string, string>();

                log("Querying movie DBs for " + title + "... ", false);
                Dictionary<string, object> imdbResult = restGet(this.imdbAPI + HttpUtility.UrlEncode(title));
                if (imdbResult != null && imdbResult["Response"].ToString() != "False")
                {
                    SQLQueryValues["imdbID"] = imdbResult["imdbID"].ToString();
                    SQLQueryValues["imdbTitle"] = imdbResult["Title"].ToString();
                    SQLQueryValues["imdbYear"] = imdbResult["Year"].ToString();
                    SQLQueryValues["imdbGenre"] = imdbResult["Genre"].ToString();
                    SQLQueryValues["imdbRating"] = imdbResult["imdbRating"].ToString();
                    SQLQueryValues["imdbRuntime"] = imdbResult["Runtime"].ToString().Replace(" min",""); //DB wants an int here and imdb returns "xxx min" so have to strip it
                    SQLQueryValues["imdbPlot"] = imdbResult["Plot"].ToString();
                }

                //Little bit of jiggling about with the data needed here because of the nested jsons (and hence nested arrays) in the return value
                Dictionary<string, object> moviedbResult = restGet(this.moviedbAPI + "?api_key=" + this.moviedbAPIkey + "&query=" + HttpUtility.UrlEncode(title)+"&page=1");
                if (moviedbResult != null && moviedbResult["total_results"].ToString() != "0") //Check if there are any results
                {
                    ArrayList results = (ArrayList)moviedbResult["results"]; //Convert the inner array of results to an arraylist
                    Dictionary<string, object> dict = (Dictionary<string, object>)results[0]; //As we know there is at least one result, take the first result and use it as a dict
                    SQLQueryValues["moviedbImage"] = dict["poster_path"].ToString(); //Have to use object in the dict to avoid implicit cast errors, so require a ToString() here
                    SQLQueryValues["moviedbID"] = dict["id"].ToString();
                }

                log(SQLQueryValues.Count+"/9 values obtained.");
                return SQLQueryValues;

            }
            /*if (type == "tv") //Not sure what we're doing with TV shows yet
            {
                Dictionary<string, string> tvdbResult = restGet(this.tvdbAPI + HttpUtility.UrlEncode(title));
            }*/
            return null;
        }

        /*
         * Makes a GET reqest to an API for a json object and returns this as a C# dictionary object so we can interrogate keys and values
         */
        private Dictionary<string, object> restGet(string url)
        {
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";//; charset=utf-8";
            httpWebRequest.Method = WebRequestMethods.Http.Get;
            try
            {
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    string json = streamReader.ReadToEnd();
                    Dictionary<string, object> jsonObj = new Dictionary<string, object>();
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    jsonObj = js.Deserialize<Dictionary<string, object>>(json);
                    return jsonObj;
                }

            }
            catch (Exception e)
            {
                log("Error when requesting data from" + url + ": " + e.Message);
                return null;
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.IsAccessible)
            {
                if (this.WindowState == FormWindowState.Minimized)
                {
                    this.ShowInTaskbar = false;
                }
            }
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            ShowInTaskbar = true;
            this.Activate();
            this.BringToFront();
            this.WindowState = FormWindowState.Normal;
        }

        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
            notifyIcon1_Click(sender, e);
        }


        
    }
}
