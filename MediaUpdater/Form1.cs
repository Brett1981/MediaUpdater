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
        private string watchedIdentifier = "";
        private bool error;
        private List<string> failedMoves = new List<string>();

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
            this.watchedIdentifier = Properties.Settings.Default.watchedIdentifier;
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
            this.log("watchedIdentifier: " + watchedIdentifier);
            this.log(">Arguments loaded.");
        }
        public Form1()
        {
            InitializeComponent();
            error = false;
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
            this.log(">Processing New Movies");
            this.processNewMovies();
            this.log(">Processing New Movies complete");
            this.log(">Deleting Old Movies");
            this.deleteOldMovies(); //Select ID and location from DB, run a directory check on the locations. If not accessible, remove entry from DB.
            this.log(">Deleting Old Movies complete");
            closeApplication();
        }

        /*
         * Gets all movies in DB and deletes the old ones that we don't need anymore due to deletion
         */
        private void deleteOldMovies()
        {
            DataTable movies = SQLSelectAll();
            int remMovies = 0;
            foreach (DataRow row in movies.Rows)
            {
               if(!Directory.Exists(@row["fileLocation"].ToString()))
               {
                   log("Movie \"" + row["imdbTitle"].ToString() + "\" not found.");
                   SQLDelete(row["movieID"].ToString());
                   remMovies++;
               }

            }
            log("Movies Deleted from DB: "+remMovies);
        }


        /*
         * Process movies into DB
         */
        private void processNewMovies()
        {
            string[] folders = System.IO.Directory.GetDirectories(@moviePath);
            string[] newMovies = new string[folders.Length]; //Get the directories and create an array ready for new movies, at least the size of the DIRs in the movies folder (in case this is a fresh run)
            int newMovieCount = 0;
            foreach (string folder in folders) //Iterate over all the movies in the folder
            {
                try
                {
                    //THIS IS THE OLD METHOD WHERE I USED A IDENT FILE FOR MOVIES. NEW METHOD USES DB ONLY.
                    /*
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
                     * */
                    if (!SQLDoesPathExist(folder) && !folder.StartsWith(moviePath+"!")) //if file not found then movie must be new and we'll add it to the array for processing
                    {                        
                        newMovies[newMovieCount] = Path.GetFileName(folder);
                        newMovieCount++;
                        this.log(folder + " is a new movie");
                    }
                }
                catch (Exception e)
                {
                    this.log(@moviePath + folder + " generated exception: " + e.Message);
                    error = true;
                }
            }
            this.log("New movies detected: "+newMovieCount);
            if (newMovieCount > 0)
            {
                foreach (string movieTitle in newMovies) //Got our new movies, time to process them.
                {
                    if (movieTitle == null) //Skip null entries (which will be at the end of the array if there are already processed films)
                        continue;

                    string watched = "0";
                    /*if(movieTitle.EndsWith(watchedIdentifier))
                        watched = "1";*/

                    //Dictionary<string, string> SQLQueryVals = checkAPI(stripWatchedIdentifier(movieTitle), "movie"); //Grab the details we need for the DB
                    Dictionary<string, string> SQLQueryVals = checkAPI(movieTitle, "movie"); //Grab the details we need for the DB
                    if (SQLQueryVals == null || SQLQueryVals.Count < 9)
                        continue;

                    string query = "INSERT INTO Movies ([imdbID], [imdbTitle], [imdbYear], [imdbGenre], [imdbRating], " +
                    "[imdbRuntime], [imdbPlot], [moviedbImage], [moviedbID], [fileLocation], [watched]) VALUES (@iID, @iTitle, @iYear, @iGenre, @iRating, @iRunTime, @iPlot, @mdbImage, @mdbID, @fileLoc, @watched);";


                    SqlCommand sqlquery = new SqlCommand(query); //Insert the query params into the query. This protects against injection attack.
                    sqlquery.Parameters.AddWithValue("@iID", SQLQueryVals["imdbID"]);
                    sqlquery.Parameters.AddWithValue("@iTitle", SQLQueryVals["imdbTitle"]);
                    sqlquery.Parameters.AddWithValue("@iYear", SQLQueryVals["imdbYear"]);
                    sqlquery.Parameters.AddWithValue("@iGenre", SQLQueryVals["imdbGenre"]);
                    sqlquery.Parameters.AddWithValue("@iRating", SQLQueryVals["imdbRating"]);
                    sqlquery.Parameters.AddWithValue("@iRunTime", SQLQueryVals["imdbRuntime"]);
                    sqlquery.Parameters.AddWithValue("@iPlot", SQLQueryVals["imdbPlot"]);
                    sqlquery.Parameters.AddWithValue("@mdbImage", SQLQueryVals["moviedbImage"]);
                    sqlquery.Parameters.AddWithValue("@mdbID", SQLQueryVals["moviedbID"]);
                    sqlquery.Parameters.AddWithValue("@fileLoc", moviePath + "\\" + movieTitle);
                    sqlquery.Parameters.AddWithValue("@watched", watched);


                    if (SQLInsert(sqlquery))
                    {
                        //No longer need to do anything on the return as we are solely useing the DB; no ident files needed anymore.
                        //createImdbIdentFile(movieTitle, SQLQueryVals["imdbID"]);
                        Console.WriteLine("Inserted " + SQLQueryVals["imdbTitle"] + " into DB.");
                    }


                }
            }
            return;

        }

        /* 
         * Strips the watched identifier from the folder name for situations where we need an API search etc.
         */
        public string stripWatchedIdentifier(string str)
        {
            //If we have the "watchedidentifier" on the end of the foldername, remove it as it will stop the searches working properly
            if (str.EndsWith(this.watchedIdentifier))
            {
                return str.Substring(0, str.Length - this.watchedIdentifier.Length);
            }

            //Or otherwise do nothing
            return str;
        }

        /*
         * REDUNDANT FUNCTION
         * Creates the identifier file so we know if a movie folder has already been processed or not
         */
        private void createImdbIdentFile(string folder, string imdbID)
        {
            log("Creating IdentFile for " + folder + "... ", false);
            try
            {
                //Create the file, make sure it starts with "imdb" to identify it
                string identFile = @moviePath + "\\" + folder + "\\imdb" + imdbID;
                File.Create(identFile);
                log("Done.");
            }
            catch (Exception e)
            {
                log("Error.");
                log(e.Message);
                error = true;
            }

        }

        /*
         * Executes the give SQLCommand on the database specified in the config file         
         */
        private bool SQLInsert(SqlCommand command)
        {
            log("Inserting into database... ", false);




            using (SqlConnection connection = new SqlConnection("User ID=" + localdbUser + ";" +
                                       "Password=" + localdbPW + ";" +
                                       "Server=tcp:" + localdbHost + ",1433;" +
                                       "Trusted_Connection=false;" + //This has to be false, otherwise connection will attempt to use Windows account credentials.
                                       "Database=" + localdbName + "; " +
                                       "connection timeout=30"))
            {
                try
                {
                    command.Connection = connection;
                    connection.Open();
                    command.ExecuteNonQuery();
                    log("Done.");
                    connection.Close();
                    return true;
                }

                catch (Exception e)
                {
                    log("Failed!");
                    log("Error with SQL insert:" + e.Message);
                    error = true;
                    connection.Close();
                    return false;
                }
            }
           
        }

        public DataTable SQLSelectAll()
        {
            log("Selecting all movies from DB... ", false);
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection("User ID=" + localdbUser + ";" +
                                           "Password=" + localdbPW + ";" +
                                           "Server=tcp:" + localdbHost + ",1433;" +
                                           "Trusted_Connection=false;" + //This has to be false, otherwise connection will attempt to use Windows account credentials.
                                           "Database=" + localdbName + "; " +
                                           "connection timeout=30"))
            {
                SqlCommand command = new SqlCommand("SELECT * FROM dbo.Movies");
                try
                {
                    command.Connection = connection;
                    connection.Open();
                    using (SqlDataReader dr = command.ExecuteReader())
                    {
                        dt.Load(dr);
                    }

                }
                catch (Exception e)
                {
                    log("Failed!");
                    log("Error with SQL Select:" + e.Message);
                    error = true;
                }
                finally { connection.Close(); }
            }
            log("Done.");
            return dt;
        }

        public bool SQLDoesPathExist(string path)
        {
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection("User ID=" + localdbUser + ";" +
                                           "Password=" + localdbPW + ";" +
                                           "Server=tcp:" + localdbHost + ",1433;" +
                                           "Trusted_Connection=false;" + //This has to be false, otherwise connection will attempt to use Windows account credentials.
                                           "Database=" + localdbName + "; " +
                                           "connection timeout=30"))
            {
                string query = "SELECT COUNT(1) FROM dbo.Movies WHERE fileLocation = @path;";
                SqlCommand command = new SqlCommand(query); //Insert the query params into the query. This protects against injection attack.
                //command.Parameters.Add("@path", SqlDbType.NVarChar).Value = path; 
                command.Parameters.AddWithValue("@path", path.Replace("\\","\\\\"));
                try
                {
                    command.Connection = connection;
                    connection.Open();

                    int movieCount = (int)command.ExecuteScalar();
                    if (movieCount > 0)
                        return true;
                    else
                        return false;

                }
                catch (Exception e)
                {
                    log("Failed to check path for " + path);
                    log("Error with SQL Select:" + e.Message);
                    error = true;
                }
                finally { connection.Close(); }
            }
            return false;
        }

        public void SQLDelete(string id)
        {
            log("Deleting MovieID "+id+"... ", false);
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection("User ID=" + localdbUser + ";" +
                                           "Password=" + localdbPW + ";" +
                                           "Server=tcp:" + localdbHost + ",1433;" +
                                           "Trusted_Connection=false;" + //This has to be false, otherwise connection will attempt to use Windows account credentials.
                                           "Database=" + localdbName + "; " +
                                           "connection timeout=30"))
            {
                string query = "DELETE FROM Movies WHERE movieID=@mID;";
                SqlCommand command = new SqlCommand(query); //Insert the query params into the query. This protects against injection attack.
                command.Parameters.AddWithValue("@mID", id);
                try
                {
                    command.Connection = connection;
                    connection.Open();
                    command.ExecuteNonQuery();

                }
                catch (Exception e)
                {
                    log("Failed!");
                    log("Error with SQL Delete:" + e.Message);
                    error = true;
                }
                finally { connection.Close(); }
            }
            log("Done.");
        }

        /*
         * Takes the response from the GET and processes these into the values we need for our SQL query. 
         */
        public Dictionary<string, string> checkAPI(string title, string type)
        {
            if (type == "movie")
            {

                var SQLQueryValues = new Dictionary<string, string>();

                //Prepopulating due to errors caused if this ends up as null
                SQLQueryValues["imdbYear"] = ""; 
                SQLQueryValues["imdbTitle"] = "";
                SQLQueryValues["imdbYear"] = "";
                SQLQueryValues["imdbGenre"] = "";
                SQLQueryValues["imdbRating"] = "";
                SQLQueryValues["imdbRuntime"] = "";
                SQLQueryValues["imdbPlot"] = "";


                log("Querying movie DBs for " + title + "... ", false);

                //Try the title first for the movie so we can grab the imdbID
                HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(this.imdbAPI + "?t=" + HttpUtility.UrlEncode(title));
                httpWebRequest.ContentType = "application/json";//; charset=utf-8";
                httpWebRequest.Method = WebRequestMethods.Http.Get;
                httpWebRequest.Timeout = 10000;
            try
            {
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                string getIDString = "";
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    getIDString = streamReader.ReadToEnd();
                }
                JavaScriptSerializer js = new JavaScriptSerializer();
                
                string imdbIDforSearch = "";
                Dictionary<string, string> getIDfromTitle = js.Deserialize<Dictionary<string, string>>(getIDString);
                if (getIDfromTitle != null && getIDfromTitle.ContainsKey("Response") && getIDfromTitle["Response"] == "True")
                {
                    imdbIDforSearch = getIDfromTitle["imdbID"];
                }
                else
                {
                    httpWebRequest = (HttpWebRequest)WebRequest.Create(this.imdbAPI + "?s=" + HttpUtility.UrlEncode(title));
                    httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                    getIDString = "";
                    
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        getIDString = streamReader.ReadToEnd();
                    }
                    Dictionary<string, Dictionary<string, string>[]> getIDfromSearch = js.Deserialize<Dictionary<string, Dictionary<string, string>[]>>(getIDString);
                    if (getIDfromSearch != null && !getIDfromSearch.ContainsKey("Response") && (getIDfromSearch["Search"][0]["imdbID"] != "" || getIDfromSearch["Search"][0]["imdbID"] != "N/A"))
                    {
                        imdbIDforSearch = getIDfromSearch["Search"][0]["imdbID"];
                    }
                }
                if (imdbIDforSearch != "") //We only get "Response" if there is a failure to find. Otherwise we get a "Search"
                {

                    //Got the ID, now do a search by the id
                    Dictionary<string, object> imdbResult = restGet(this.imdbAPI + "?i=" + HttpUtility.UrlEncode(imdbIDforSearch));
                    if (imdbResult != null && imdbResult["Response"].ToString() != "False")
                    {

                        //Using inline if's: result = (<condition> ? <true> : <false>);
                        SQLQueryValues["imdbID"] = (imdbResult.ContainsKey("imdbID") ? imdbResult["imdbID"].ToString() : "");
                        SQLQueryValues["imdbTitle"] = (imdbResult.ContainsKey("Title") ? imdbResult["Title"].ToString() : "");
                        SQLQueryValues["imdbYear"] = (imdbResult.ContainsKey("Year") ? imdbResult["Year"].ToString() : "");
                        SQLQueryValues["imdbGenre"] = (imdbResult.ContainsKey("Genre") ? imdbResult["Genre"].ToString() : "");
                        SQLQueryValues["imdbRating"] = (imdbResult.ContainsKey("imdbRating") ? imdbResult["imdbRating"].ToString() : "");
                        SQLQueryValues["imdbRuntime"] = (imdbResult.ContainsKey("Runtime") ? imdbResult["Runtime"].ToString() : "");
                        SQLQueryValues["imdbPlot"] = (imdbResult.ContainsKey("Plot") ? imdbResult["Plot"].ToString() : "");
                    }
                }
            }
            catch (Exception e)
            {
                log("Error getting IMDB data: " + e.Message);
            }
                

                //Little bit of jiggling about with the data needed here because of the nested jsons (and hence nested arrays) in the return value
                //If we have an IMDB ID use that, as it's exact.
                Dictionary<string, object> moviedbResult = null;

                if (SQLQueryValues.ContainsKey("imdbID") && SQLQueryValues["imdbID"].Length > 3)
                {
                    moviedbResult = restGet(this.moviedbAPI + "find/" + SQLQueryValues["imdbID"] + "/?api_key=" + this.moviedbAPIkey + "&external_source=imdb_id");
                }

                //If we didn't get anything, we're going to have to search for the details. We'll also use the SQLQueryValues["imdbYear"] result above to narrow down the search results from the moviedbapi if available
                if (moviedbResult != null && !moviedbResult.ContainsKey("moviedbID")) 
                {
                    if (SQLQueryValues["imdbYear"] == "" || SQLQueryValues["imdbYear"] == "N/A")
                        moviedbResult = restGet(this.moviedbAPI + "search/movie?api_key=" + this.moviedbAPIkey + "&query=" + HttpUtility.UrlEncode(title) + "&page=1");
                    else
                        moviedbResult = restGet(this.moviedbAPI + "search/movie?api_key=" + this.moviedbAPIkey + "&query=" + HttpUtility.UrlEncode(title) + "&page=1&year=" + SQLQueryValues["imdbYear"]);
                }

                if (moviedbResult != null && moviedbResult.ContainsKey("movie_results")) //Check if there are any results
                {
                    ArrayList results = (ArrayList)moviedbResult["movie_results"]; //Convert the inner array of results to an arraylist

                    if (results.Count > 0) //Sometimes moviedb can give results > 0 but an empty array
                    {
                        Dictionary<string, object> dict = (Dictionary<string, object>)results[0]; //Take the first result and use it as a dict
                        SQLQueryValues["moviedbImage"] = (dict["poster_path"] != null ? dict["poster_path"].ToString() : ""); //Have to use object in the dict to avoid implicit cast errors, so require a ToString() here
                        SQLQueryValues["moviedbID"] = (dict["id"] != null ? dict["id"].ToString() : "");
                    }

                }
                else
                {
                    if (moviedbResult != null && moviedbResult["total_results"].ToString() != "0")
                    {
                        ArrayList results = (ArrayList)moviedbResult["results"]; //Convert the inner array of results to an arraylist

                        if (results.Count > 0) //Sometimes moviedb can give results > 0 but an empty array
                        {
                            Dictionary<string, object> dict = (Dictionary<string, object>)results[0]; //Take the first result and use it as a dict
                            SQLQueryValues["moviedbImage"] = (dict["poster_path"] != null ? dict["poster_path"].ToString() : ""); //Have to use object in the dict to avoid implicit cast errors, so require a ToString() here
                            SQLQueryValues["moviedbID"] = (dict["id"] != null ? dict["id"].ToString() : "");
                        }

                    }
                }
                if (SQLQueryValues.Count < 9)
                {
                    error = true;
                    failedMoves.Add(title);
                    Console.WriteLine("Could not insert " + title+". "+SQLQueryValues.Count + "/9 values obtained from API calls. Skipping DB entries.");
                    log(SQLQueryValues.Count + "/9 values obtained.Unable to obtain all values from API calls for \"" + title + "\". Skipping DB entries.");
                    log("imdbID: " + (SQLQueryValues.ContainsKey("imdbID") ? SQLQueryValues["imdbID"].ToString() : "EMPTY"));
                    log("imdbTitle: " + (SQLQueryValues.ContainsKey("imdbTitle") ? SQLQueryValues["imdbTitle"].ToString() : "EMPTY"));
                    log("imdbYear: " + (SQLQueryValues.ContainsKey("imdbYear") ? SQLQueryValues["imdbYear"].ToString() : "EMPTY"));
                    log("imdbGenre: " + (SQLQueryValues.ContainsKey("imdbGenre") ? SQLQueryValues["imdbGenre"].ToString() : "EMPTY"));
                    log("imdbRating: " + (SQLQueryValues.ContainsKey("imdbRating") ? SQLQueryValues["imdbRating"].ToString() : "EMPTY"));
                    log("imdbRuntime: " + (SQLQueryValues.ContainsKey("imdbRuntime") ? SQLQueryValues["imdbRuntime"].ToString() : "EMPTY"));
                    log("imdbPlot: " + (SQLQueryValues.ContainsKey("imdbPlot") ? SQLQueryValues["imdbPlot"].ToString() : "EMPTY"));
                    log("moviedbImage: " + (SQLQueryValues.ContainsKey("moviedbImage") ? SQLQueryValues["moviedbImage"].ToString() : "EMPTY"));
                    log("moviedbID: " + (SQLQueryValues.ContainsKey("moviedbID") ? SQLQueryValues["moviedbID"].ToString() : "EMPTY"));

                }
                else
                {
                    log(SQLQueryValues.Count + "/9 values obtained.");
                }
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
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";//; charset=utf-8";
            httpWebRequest.Method = WebRequestMethods.Http.Get;
            httpWebRequest.Timeout = 10000;
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
                error = true;
                return null;
            }
        }

        //On Form Resize, if it's minimised them hide the taskbar icon.
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

        //If we click on the systemtray icon, restore the window
        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            ShowInTaskbar = true;
            this.Activate();
            this.BringToFront();
            this.WindowState = FormWindowState.Normal;
        }

        //If we click on the balloon tooltip, perform the same action as if we had clicked on the systemtray icon
        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
            notifyIcon1_Click(sender, e);
        }

        private void closeApplication()
        {
            Console.WriteLine("Processing complete.");
            if (error)
            {
                log("Writing errors.");
                File.WriteAllText(@Application.StartupPath + "\\Logs\\ErrorLog." + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt", output.Text);
                Console.WriteLine("Errors encountered, please refer to application log");
            }
            if (failedMoves.Count > 0)
            {
                log("Writing failed movies.");
                File.WriteAllLines(@Application.StartupPath + "\\Logs\\FailedMoves." + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt", failedMoves);
                Console.WriteLine("Movie(s) did not update into DB. Please see failed movie log (" + DateTime.Now.ToString("yyyyMMdd-HHmmss")+")");
            }
            log("Closing application.");
            Application.Exit();
        }
        
    }
}
