angular.module('myApp', [])

    .controller('Main', ['$http', "$scope", function ($http, $scope) {

        var token = localStorage.getItem('jstalks_token');

        $scope.isLoggedIn = !!token;
        $scope.menuItems = [
            { name: "Home", url: '', selected: true, icon: "home" },
            { name: "Azure AD", url: 'htm/azureAD.htm', icon: "group" },
            { name: "Exchange", url: 'htm/exchange.htm', icon: "mail" },
            { name: "Calendar", url: 'htm/calendar.htm', icon: "calendar" },
            { name: "OneDrive", url: 'htm/onedrive.htm', icon: "briefcase" },
            { name: "Onenote", url: 'htm/onenote.htm', icon: "notebook" },
            { name: "SharePoint", url: 'htm/sharepoint.htm', icon: "work", isRedirect: true },
            { name: "Yammer", url: 'htm/yammer.htm', icon: "org", isRedirect: true }
        ];

        if ($scope.isLoggedIn) {

            $http.defaults.headers.common.Authorization = "Bearer " + token;
            console.log("logged in");
            //$scope.selectedTemplate = "htm/calendar.htm";
        }

        $scope.login = function () {
            var clientId = 'client_id';
            var redirectUrl = 'redirect_url';
            var resource = "https://graph.microsoft.com/";
            window.location = 'https://login.windows.net/common/oauth2/authorize?response_type=token&client_id=' + clientId + '&resource=' + resource + '&redirect_uri=' + redirectUrl;
        };

        $scope.logout = function () {
            localStorage.removeItem('jstalks_token');
            window.location.reload();
        }

        $scope.menuClick = function (menuItem) {

            angular.forEach($scope.menuItems, function (value, key) {
                value.selected = false;
            });

            menuItem.selected = true;

            if (menuItem.isRedirect)
                window.location = menuItem.url;
            else
                $scope.selectedTemplate = menuItem.url;

        };

    }])

	  
	  .controller('AzureAD', ['$http', "$scope", function($http, $scope) {
		
		var token = localStorage.getItem('jstalks_token');
		
		if (token) {
		    $http.get("https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,jobTitle,assignedLicenses,officeLocation,country,mobilePhone,givenName,surname")
				.success(function successCallback(response) {
				//$scope.data = JSON.stringify(response, "", 2);
				    $scope.users = response.value.filter(function (user) {
				        return user.assignedLicenses && user.assignedLicenses.length && user.jobTitle;
				    });

				    angular.forEach($scope.users, function (user) {
				        loadImage(user);
				    });

			}).error(function (errorObj, errorCode) { 
				$scope.error = JSON.stringify(errorObj);
				$scope.errorCode = errorCode;
			});
		}

		function loadImage (user) {
			$http({
			  url: 'https://graph.microsoft.com/v1.0/users/' + user.mail + '/photo/$value',
			  method: 'GET',
			  responseType: 'blob'
			}).success(function(response) {
				var url = window.URL || window.webkitURL;
				user.image = url.createObjectURL(response);
			});
		}
	}])
	
	.controller('OneDrive', ['$http', "$scope", function($http, $scope) {

        var token = localStorage.getItem('jstalks_token');
		var rootUrl = "https://graph.microsoft.com/v1.0/me/drive/root/children?$select=name,size,lastModifiedDateTime,contentUrl,id,file,folder";
		
		$scope.isLoggedIn = !!token;
		
        if ($scope.isLoggedIn) {

			$http.defaults.headers.common.Authorization = "Bearer " + token;
			openFolder("", "root");
        }
		
		$scope.openFile = function(file) {
			window.open(file.webUrl);
		};
		
		$scope.openFolder = openFolder;

		function openFolder(folderId, parent) {
			
			$scope.files = [];
			$scope.data = "";
			
			if (folderId === "" || folderId === "root")
				var url = rootUrl;
			else 
				var url = "https://graph.microsoft.com/v1.0/me/drive/items/" + folderId + "/children?$select=name,size,lastModifiedDateTime,contentUrl,id,file,folder,webUrl";
							
			$http.get(url)
			  .success(function (response) {
				$scope.files = response.value;
				//$scope.data = JSON.stringify(response, "", 2);
				$scope.parent = parent;
			  }).error(function (errorObj, errorCode) {
			      $scope.error = JSON.stringify(errorObj);
			      $scope.errorCode = errorCode;
			  });;
		}
		
	}])
	
	.controller('Exchange', ['$http', "$scope", "$sce", function($http, $scope, $sce) {

        var token = localStorage.getItem('jstalks_token');
        var rootUrl = "https://graph.microsoft.com/v1.0/me/messages?$select=subject,ccRecipients,receivedDateTime,from,hasAttachments,id,isRead,toRecipients,bodyPreview";

		$scope.isLoggedIn = !!token;

		$http.defaults.headers.common.Authorization = "Bearer " + token;
		
        if ($scope.isLoggedIn) {
            loadNextPage(rootUrl);
		}
		
        $scope.openMessage = function (message) {

            $scope.singleMessage = { body: "Loading ..." };

            $http.get("https://graph.microsoft.com/v1.0/me/messages/" + message.id + "?$select=body")
                .success(function (response) {
                    $scope.singleMessage = response;
                    $scope.singleMessage.body.content = $sce.trustAsHtml($scope.singleMessage.body.content);
                });
        };

        $scope.backToList = function () {
            $scope.singleMessage = null;
        };

        $scope.composeCancelClick = function () {
            $scope.composeOpen = false;
        };

        $scope.composeSaveClick = function(subject, to, body) {
			
            $scope.composeOpen = false;
            $scope.composeTo = to;

			var body = {
			    "Subject": subject,
				"Body": {
					"ContentType": "HTML",
					"Content": body
				},
				"ToRecipients": [
					{
						"EmailAddress": {
							"Address": to
						}
					}
				]
			};

			$http.post("https://graph.microsoft.com/beta/me/Messages", body).then(function successCallback(response) {

				$http.post("https://graph.microsoft.com/beta/me/Messages/" + response.data.id + "/send")
					.then(function successCallback(response) {
					    $scope.data = JSON.stringify(response, undefined, 2);
					    $scope.composeBanner = true;
					});			
				
			});
			
        };

        $scope.loadMore = function () {
            loadNextPage($scope.nextLink);
        };
		
        function loadNextPage(url) {
            $http.get(url)
                .success(function (response) {
                    $scope.data = JSON.stringify(response, "", 2);

                    if (!$scope.messages)
                        $scope.messages = [];

                    $scope.messages = $scope.messages.concat(response.value);
                    $scope.nextLink = response["@odata.nextLink"];
                }).error(function (errorObj, errorCode) {
                    $scope.error = JSON.stringify(errorObj);
                    $scope.errorCode = errorCode;
                });;
        }

	}])
	
	.controller('OneNote', ['$http', "$scope", "$sce", function($http, $scope, $sce) {

        var token = localStorage.getItem('jstalks_token');
		
		$scope.isLoggedIn = !!token;
		
        if ($scope.isLoggedIn) {

			$http.defaults.headers.common.Authorization = "Bearer " + token;
			
			$http.get("https://graph.microsoft.com/beta/me/notes/pages?$select=title,id,contentUrl,lastModifiedTime").success(function (response) {

			    //$scope.data = JSON.stringify(response, "", 2);
			    $scope.pages = response.value.filter(function (page) { return page && page.title; });

			}).error(function (errorObj, errorCode) {
			    $scope.error = JSON.stringify(errorObj);
			    $scope.errorCode = errorCode;
			});
			
        }
		else {
			$scope.error = "Not logged in";
		}
		
		$scope.openPage = function(page) {

		    $scope.pageContent = $sce.trustAsHtml("Loading ...");

			$http.get(page.contentUrl).success(function (response) {
			  
				$scope.data = "";
				$scope.pageContent = $sce.trustAsHtml(response);
				
			})
			
		};

		$scope.back = function () {
		    $scope.pageContent = null;
		};

	}])

    .controller('Calendar', ['$http', "$scope", function ($http, $scope) {

        var token = localStorage.getItem('jstalks_token');

        $scope.isLoggedIn = !!token;

        if ($scope.isLoggedIn) {

            $http.defaults.headers.common.Authorization = "Bearer " + token;
            loadData("https://graph.microsoft.com/v1.0/me/events?$select=subject,bodyPreview,id,start,attendees,organizer");

        }
        else {
            $scope.error = "Not logged in";
        }

        $scope.loadMore = function () {
            loadData($scope.nextLink);
        };

        function loadData(url) {
            $http.get(url).success(function (response) {

                //$scope.data = JSON.stringify(response, "", 2);

                if (!$scope.events)
                    $scope.events = [];

                $scope.events = $scope.events.concat(response.value);

                $scope.nextLink = response["@odata.nextLink"];

            }).error(function (errorObj, errorCode) {
                $scope.error = JSON.stringify(errorObj);
                $scope.errorCode = errorCode;
            });
        }

    }])
	
    .filter("formatDate", function() {
	
        var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul","Aug", "Sep", "Oct", "Nov", "Dec" ];

        return function (dateString) {

            var d = new Date(dateString);

            return d.getDate() + " " + monthNames[d.getMonth()] + " " + d.getFullYear() + " " +
                ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2);
        };

    })

	.filter("friendlyDate", function() {
		
			return function(date) {
				date = new Date(date);
				var seconds = Math.floor((new Date() - date) / 1000);

				var interval = Math.floor(seconds / 31536000);

				if (interval > 1) {
					return interval + " years ago";
				}
				interval = Math.floor(seconds / 2592000);
				if (interval > 1) {
					return interval + " months ago";
				}
				interval = Math.floor(seconds / 86400);
				if (interval > 1) {
					return interval + " days ago";
				}
				interval = Math.floor(seconds / 3600);
				if (interval > 1) {
					return interval + " hours ago";
				}
				interval = Math.floor(seconds / 60);
				if (interval > 1) {
					return interval + " minutes ago";
				}
				return Math.floor(seconds) + " seconds ago";
			};
		
		})
		
		.filter ("fileSize", function() {
		
			return function(sizeInBytes) {
				var kb = Math.round(sizeInBytes / 1000);
				if (kb >= 1000)
					return Math.round(kb / 1000).toString() + " MB";
				else 
					return kb.toString() + " KB";
			};
		
		})

        .filter("formatUsers", function () {

            return function (input) {

                var result = "";

                angular.forEach(input, function (value) {
                    if (value)
                        result += "<a href='mailto:" + value.emailAddress.address + "'>" + value.emailAddress.name + "</a>" + "; ";
                });

                result = result.substring(0, result.length - 2);

                return result;
            };

        })

        .filter("formatUser", function () {

            return function (value) {

                if (!value)
                    return "";

                return "<a href='mailto:" + value.emailAddress.address + "'>" + value.emailAddress.name + "</a>";

            };

        })

        .directive('toHtml', function () {
            return {
                restrict: 'A',
                link: function (scope, el, attrs) {
                    el.html(scope.$eval(attrs.toHtml));
                }
            };
        })