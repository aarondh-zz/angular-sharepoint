/*
 * angular-sharepoint version 1.0.3 2/05/2015
 *
 * Written By: Aaron G. Daisley-Harrison <aaron@daisley-harrison.com>
 *
 * angular-sharepoint is an angular module that extends 
 * the SharePoint 2013 Client Model and REST services to provide
 * generic angular controller functionality to acces SharePoint lists and managed meta data
 *
 * For more information on angular please visit http://angular.org
 */

var ngSharePoint = angular.module('angular-sharepoint', ['ngResource','ui.event']);

ngSharePoint.constant('SITE', undefined);

ngSharePoint.constant('WEB', undefined);

ngSharePoint.service('SharePoint',['SITE','WEB', function(site,web) {
    var webPath = web;
    var sitePath = site;
    return {
       escape: function( text ) {
          if ( text[text.length-1] == '/' ) {
             text = text.substr(0,text.length-1); //remove trailing slash
          }
          return text.replace(/:/g,"\\:");
       },
       getSite: function getSite() {
          if ( typeof sitePath === 'undefined' ) {
          	return _spPageContextInfo.siteAbsoluteUrl; 
          }
          else {
          	return sitePath;
          }
       },
       getSiteEscaped: function getSiteEscaped() {
          return this.escape(this.getSite()); 
       },
       getWeb: function getWeb() {
          if ( typeof webPath  === 'undefined' ) {
		        var spWeb = _spPageContextInfo.webServerRelativeUrl;
		        return this.escape(spWeb);
	        }
	        else {
		        return this.escape(webPath);
	        }
       },
       setSite: function setSite(site) {
         sitePath = site;
       },
       setWeb: function setWeb(web) {
         webPath = web;
       },
       getDisplayMode: function getDisplayMode() {
            var $displayModeName = angular.element('#MSOSPWebPartManager_DisplayModeName');
            return $displayModeName.val();
       }
    };
}]);
ngSharePoint.factory('httpRequestInterceptor', function (SharePoint) {

  return {

    request: function (config) {
      var url = config.url;
      var sep = url.indexOf('?');
      var path;
      var query;
      if ( sep < 0 ) {
         path = url;
         query = "";
      }
      else {
         path = url.substr(0,sep);
         query = url.substr(sep);
      }
      path = path.replace(/%2F/g,"/"); //reverse slash escaping in the path
      url = path + query;
      config.url = url;
      return config;

    }

  };

});

 

ngSharePoint.config(function ($httpProvider) {

  $httpProvider.interceptors.push('httpRequestInterceptor');

});


ngSharePoint.service("TermStoreService",['$q','SharePoint', function($q, sharePoint) {
  
  return {
    options: {
      lcid: undefined,          //The language used to find labels
      isReturnLabels: false,     //If true, all labels of a term are returne
      verbose: false             //if true, verbose messages are logged to the console
    },
    state: 'not ready',
    termById: {},
    deferredWorkToBeDone: [],
    _init: function() {
      var me = this;
      me.state = 'initialized';
      },
    _executeOnTaxonomyScriptLoaded: function( done ) {
  		var scriptbase = sharePoint.getSite() + "/_layouts/15/";
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () { console.log("Initiating SP.ClientContext"); });
        SP.SOD.executeOrDelayUntilScriptLoaded(function() {
	      SP.SOD.registerSod('sp.taxonomy.js', scriptbase + 'sp.taxonomy.js');
	      SP.SOD.executeFunc('sp.taxonomy.js', false, done );
	    }, 'sp.js');
	},
	_runAfterInit: function( done ) {
	   var me = this;
	   if ( me.state === 'initialized' ) {
	      done(); //already initialized
	   }
	   else {
	      this._executeOnTaxonomyScriptLoaded( function() {
	         me._init();
	         console.log('sharepoint script initialized');
	         done();           
	      });
	   }
	},
	_getLabels: function( spTerm ) {
	   var deferred = $q.defer();
       var context = SP.ClientContext.get_current();
	   var labelCollection = spTerm.get_labels();
	   context.load(labelCollection);
       context.executeQueryAsync(function(){
		   var labelEnumerator = labelCollection.getEnumerator();
		   var labels= [];
		   while (labelEnumerator.moveNext()) {
		      var label = labelEnumerator.get_current();
		      labels.push( { isDefaultForLanguage: label.get_isDefaultForLanguage(), language: label.get_language(), value: label.get_value()});
		   }
		   deferred.resolve(labels);
	   }, function(reason){
	      deferred.reject(reason);
	   });
	   return deferred.promise;
	},
	_getNavigationUrl: function( term, options) {
	    var navigationUrl = term.LocalCustomProperties._Sys_Nav_SimpleLinkUrl;
	    if ( typeof navigationUrl === 'undefined') {
		    if ( typeof options.baseNavigationUrl === 'undefined' ) {
		       navigationUrl = "";
		    }
		    else {
		       navigationUrl = options.baseNavigationUrl;
		    }
		    if ( navigationUrl[navigationUrl.length-1] !== '/') {
		      navigationUrl += '/';
		    }
		    navigationUrl += term.Name.toLowerCase().replace(/[^a-z0-9]+/,'-');
	     }
	     return navigationUrl;
	},
	_createTermFromSPTerm: function( spTerm, options ) {
	 	var term = angular.extend(
	       {},
	       spTerm.get_objectData().get_properties(),
		   {
		     Id: spTerm.get_id().toString()
		   }
		);
		if ( typeof term.LocalCustomProperties === 'undefined' ) {
		    term.LocalCustomProperties = {};
		}
		term.LocalCustomProperties.NavigationUrl = this._getNavigationUrl( term, options );
		if ( options.isReturnLabels ) {
			this._getLabels(spTerm).then( function(labels){
			   term.labels = labels;
			});
		}
		return term;
	},
	_processLoadedTerms: function(termEnumerator, options) {
		var terms = [];
		while(termEnumerator.moveNext()){
			var term = this._createTermFromSPTerm( termEnumerator.get_current(), options );
			terms.push(term);
			var deferred = this.termById[term.Id];
			if ( typeof deferred === 'undefined' ) {
			   deferred = this.termById[term.Id] = $q.defer();
			}
		    if ( options.verbose ) {
		      console.log('resolved ' + term.Id + " -> " + term.Name);
		    }
		   deferred.resolve(term);
		}
		return terms;
	},
    _loadTermsById: function( termIds, done, fail ) {
      var me = this;
      var context = SP.ClientContext.get_current();
      var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
      var terms = session.getTermsById(termIds);
	  console.log('requesting ' + termIds.length + ' terms');
 	  context.load(terms);
      context.executeQueryAsync(function(){
         try {
            var termEnumerator = terms.getEnumerator();
 	        console.log('terms received');
            me._processLoadedTerms(termEnumerator, me.options);
		    if ( typeof done === 'function') {
		        done();
		    } 
		 }
		 catch(err) {
			console.log(err);
		    if ( typeof fail === 'function') {
		        fail();
		    } 
		 }
	  },function(sender,args){
		console.log(args.get_message());
	    if ( typeof fail === 'function') {
	        fail();
	    } 
	  });
    },
    _loadTermStore: function( termStore, options ) {
       var me = this;
       var deferred = $q.defer();
       if ( typeof termStore=== 'object' ) {
          deferred.resolve(termStore);
          return deferred.promise;
       }
 	   this._runAfterInit( function() {
	      var context = SP.ClientContext.get_current();
	      var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
	      if ( typeof termStore === 'undefined' ) {
	          console.log('requesting the default site collection term store.');
		      termStore = session.getDefaultSiteCollectionTermStore();
		      context.load(termStore);
		      context.executeQueryAsync(function(){
			  me.termStores = [];
		         try
		         {
		            me.termStores.push(termStore);
					console.log( 'Default site colleciton term store "' + termStore.get_name() + '" was loaded.');
		            deferred.resolve(termStore);
		         }
		         catch(err) {
					console.log(err);
				    deferred.reject( err );
		         }
		      });
	      }
	      else if ( typeof termStore=== 'string' ) {
	          var termStoreName = termStore;
	          console.log('requesting termStore "' + termStoreName + '"');
		      var termStores = session.get_termStores();
		      context.load(termStores);
		      context.executeQueryAsync(function(){
	          var matchingTermStore = null;
				 try {
				    var termStoreEnumerator = termStores.getEnumerator();
					me.termStores = [];
					while(termStoreEnumerator.moveNext()){
					  termStore = termStoreEnumerator.get_current();
					  me.termStores.push( termStore );
					  if ( termStore.get_name() === termstoreOrTermStoreName ) {
				         console.log( 'Term store "' + termStoreName + '" was loaded.');
					     matchingTermStore = termStore;
					  }
					}
					if ( matchingTermStore === null ) {
				       deferred.reject( 'Term store "' + termStoreName + '" was not found.');
					}
					else {
				       deferred.resolve( matchingTermStore );
					}
				 }
				 catch(err) {
					console.log(err);
				    deferred.reject( err );
				 }
		      });
	      }
	      else {
			  deferred.reject( "parameter termStore not a term store or a term store name" );
	      }
 	   });
       return deferred.promise;
    },
    _loadTermSet: function( termStore, termSetName, options ) {
		var me = this;
		var deferred = $q.defer();
        options = angular.extend({}, me.options, options);
		console.log('requesting term set "' + termSetName + '"');
        var context = SP.ClientContext.get_current();
        var lcid = typeof options.lcid == 'undefined' ? termStore.get_defaultLanguage() : options.lcid;
		var termSetCollection = termStore.getTermSetsByName(termSetName, lcid );
		context.load(termSetCollection);
		context.executeQueryAsync(function(){
		    var termSetEnumerator = termSetCollection.getEnumerator();
		    if ( termSetEnumerator.moveNext() ) {
		        var termSet = termSetEnumerator.get_current();
				console.log( 'Term set "' + termSet.get_name() + '" was loaded.');
		        deferred.resolve(termSet);
			}
			else {
		       var err = 'Term set "'+ termSetName + '" was not found.';
			   console.log(err );
			   deferred.reject(err );
			}
		},function(sender,args){
	       var err = 'Error loading term set "'+ termSetName + '": ' + args.get_message();
		   console.log(err );
		   deferred.reject(err );
	    });
		return deferred.promise;
    },
    _loadGroup: function( termStore, groupName, options ) {
		var me = this;
		var deferred = $q.defer();
		console.log('requesting term groups from term set "' + termSet.get_name() + '"');
        var context = SP.ClientContext.get_current();
		var groupCollection = termStore.groups;
		context.load(groupCollection);
		context.executeQueryAsync(function(){
		    var groupEnumerator = groupCollection.getEnumerator();
		    if ( typeof groupName === 'undefined' ) {
		       var groups = [];
			   while ( groupEnumerator.moveNext() ) {
			     groups.push(group);
			   }
			   console.log( 'Groups for term set "' + termSet.get_name() + '" were loaded.');
			   deferred.resolve(groups);
		    }
		    else {
			    while ( groupEnumerator.moveNext() ) {
			        var group  = groupEnumerator.get_current();
			        if ( group.get_name() === groupName ) {
						console.log( 'Term group "' + group.get_name() + '" was loaded.');
				        deferred.resolve(group);
				        return;
			        }
				}
		       var err = 'Term group "'+ groupName + '" was not found.';
			   console.log(err );
			   deferred.reject(err );
		    }
		},function(sender,args){
	       var err = 'Error loading term set "'+ termSetName + '": ' + args.get_message();
		   console.log(err );
		   deferred.reject(err );
	    });
		return deferred.promise;
    },
    _loadTermsFromPath: function( container, path, options ) {
        var me = this;
        var deferred = $q.defer();
        var sep = path.indexOf(':');
        if ( sep < 0 ) {
           component = path.trim();
           path = '';
        }
        else {
           component = path.substr(0,sep).trim();
           path = path.substr(sep+1);
        }
        if ( component.length === 0 ) {
           deferred.resolve(container);
           return;
        }
	    var termCollection = container.get_terms();
        var context = SP.ClientContext.get_current();
		context.load(termCollection);
        if ( options.verbose ) {
			console.log('requesting terms from "' + container.get_name() + '"');
        }
		context.executeQueryAsync(function(){
			 try {
			    var isFound = false;
			    var termEnumerator = termCollection.getEnumerator();
			    var terms = me._processLoadedTerms(termEnumerator, options);
			    if ( component === '*' ) {
			       isFound = true;
			       deferred.resolve( terms );
			       if ( path.length !== 0 ) {
			           if ( path[0] !== "*" ) {
		                  deferred.reject( "Path components following a astrisk (*) can only be additional astrisks to denote depth e.g. *:*:*" );
			           }
			           var i = 0;
				       termEnumerator = termCollection.getEnumerator();
				       while( termEnumerator.moveNext() ) {
				         var term = terms[i++];
				         if ( term.TermsCount > 0 ) {
				         var spTerm = termEnumerator.get_current();
				             var childOptions = angular.extend({}, options, { baseNavigationUrl: term.NavigationUrl} );
			                 me._loadTermsFromPath( spTerm, path, childOptions ).then( Function.createDelegate(term ,function(children){
			                   this.terms = children;
			                 }), function(reason) {
			                    deferred.reject( reason );
			                 });
		                 }
				       }
			       }
			    }
			    else if ( path.length === 0 ) {
			       for( i in terms ) {
			         var term = terms[i];
			         if ( term.Name == component ) {
			            isFound = true;
		                deferred.resolve(term);
		                break;
			         }
			       }
			    }
			    else {
			       var i = 0;
			       termEnumerator = termCollection.getEnumerator();
			       while( termEnumerator.moveNext() ) {
			         var term = terms[i++];
			         var spTerm = termEnumerator.get_current();
			         if ( spTerm.get_name() === component ) {
			            isFound = true;
				        var childOptions = angular.extend({}, options, { baseNavigationUrl: term.NavigationUrl} );
		                me._loadTermsFromPath( spTerm, path, childOptions ).then(function(result){
		                   deferred.resolve(result);
		                }, function(reason) {
		                  deferred.reject( reason );
		                });
		                break;
			         }
			       }
			    }
			    if ( isFound ) {
		            if ( options.verbose ) {
	                  console.log('child term "' + component + '" found in "' + container.get_name() + '"');
		            }
			    }
			    else {
		           throw 'child term "' + component + '" was not found in "' + container.get_name() + '"';
			    }
			 }
			 catch(err) {
				console.log(err);
			    deferred.reject( err );
			 }
		},function(sender,args){
           var err = 'Error loading terms from term set "'+ termSetName + '": ' + args.get_message();
		   console.log(err);
		   deferred.reject(err);
		});
		return deferred.promise;
    },
    _loadTermsFromTermSet: function( termSet, options ) {
        var me = this;
        var deferred = $q.defer();
	    var termCollection = termSet.get_terms();
        var context = SP.ClientContext.get_current();
		context.load(termCollection);
		console.log('requesting terms from term set "' + termSet.get_name() + '"');
		context.executeQueryAsync(function(){
			 try {
			    var termEnumerator = termCollection.getEnumerator();
			    var terms = me._processLoadedTerms(termEnumerator, options );
		        console.log('loaded '+ terms.length + ' terms from term set "' + termSet.get_name() + '"');
			    deferred.resolve( terms );
			 }
			 catch(err) {
				console.log(err);
			    deferred.reject( err );
			 }
		},function(sender,args){
           var err = 'Error loading terms from term set "'+ termSetName + '": ' + args.get_message();
		   console.log(err);
		   deferred.reject(err);
		});
		return deferred.promise;
    },
    _loadTermsFromTermStoreTermSet: function( termStore, termSetName, options ) {
		var me = this;
		var deferred = $q.defer();
		this._loadTermStore(termStore, options).then(function(termStore) {
		   me._loadTermSet(termStore, termSetName, options).then( function(termSet) {
		         me._loadTermsFromTermSet(termSet, options).then(function(terms){
		            deferred.resolve( {termStore: termStore, termSet: termSet, terms: terms});
		         }, function(reason) {
		           deferred .reject(reason);
		         });
		   }, function(reason) {
		      deferred .reject(reason);
		   });
		}, function(reason) {
		  deferred .reject(reason);
		});
		return deferred.promise;
    },
    loadTerms: function( path, options ) {
       var me = this;
       var deferred = $q.defer();
       if ( typeof options === 'undefined') {
          options = {};
       }
       options = angular.extend({}, me.options, options);
       if ( typeof options.termStore === 'undefined' ) {
	       if ( typeof path === 'string' && options.termStoreFromPath === true ) {
	          var sep = path.indexOf(':');
	          if ( sep >= 0 ) {
	            options.termStore = path.substr(0,sep);
	            path = path.substr(sep);
	          }
	          else {
	            options.termStore = path;
	            path = "*";
	          }
	       }	       
       }
       this._loadTermStore(options.termStore, options).then(function(termStore) {
	       if ( typeof path === 'string' ) {
	          var sep = path.indexOf(':');
	          if ( sep < 0 ) {
	             me._loadTermsFromTermStoreTermSet(termStore, path, options).then(function(result) {
   	                deferred.resolve(result.terms);
	             }, function(reason){
   	                deferred.reject(reason);
	             });
	          }
	          else {
	             var termSetName = path.substr(0,sep);
	             path = path.substr(sep+1);
	             me._loadTermSet(termStore, termSetName, options).then(function(termSet) {
		             me._loadTermsFromPath(termSet, path, options).then(function(result) {
	   	                deferred.resolve(result);
		             }, function(reason){
	   	                deferred.reject(reason);
		             });
				 }, function(reason){
	   	            deferred.reject(reason);
		         });
	          }
	       }
	       else {
	       	   deferred.resolve([]);
	   	   }
	   	   }, function(reason) {
   	      deferred.reject(reason);
   	   });
       return deferred.promise;
    },
    resolveTerm: function(termGuid) {
       var deferred = this.termById[termGuid];
       if ( typeof deferred === 'undefined' ) {
          deferred = $q.defer();
          deferred._termId = termGuid;
          this.termById[termGuid] = deferred;
          this.deferredWorkToBeDone.push( deferred );
       }
       return deferred.promise;
    },
    resolveQueue: [],
    _resolveFromQueue: function() {
      var me = this;
      me.isResolving = true;
      if ( me.resolveQueue.length > 0 ) {
          var work = me.resolveQueue.pop();
          if ( work.toBeDone.length === 0 ) {
             me._resolveFromQueue();
             if ( typeof work.done === 'function' ) {
               setTimeout(work.done,1);
             }
             return;
          }
	      this._runAfterInit( function() {
		      var termIds = [];
		      for( key in work.toBeDone ) {
		         var deferred = work.toBeDone[key];
		         termIds.push(deferred._termId);
		      }
		      me._loadTermsById( termIds, function() {
                 me._resolveFromQueue();
                 if ( typeof work.done === 'function' ) {
                   work.done();
                 }
		      }, 
		      function() {
		         //failed
		         me.isResolving = false;
		      }
		      );
	       } );
       }
       else {
         me.isResolving = false;
      }
    },
    resolveAll: function(done, force) {
      var resolveWork = { toBeDone: this.deferredWorkToBeDone, done: done};
      this.resolveQueue.push(resolveWork);
      this.deferredWorkToBeDone = [];
      if ( this.isResolving ) {
        return;
      }
      this._resolveFromQueue();
    },
    processAllTaxonomyFields: function( items, done ) {
          for(key in items  ) {
             var item = items [key];
             for( propertyName in item ) {
                var property = item[propertyName];
                if ( property && property.__metadata ) {
                   if ( property.__metadata.type == 'SP.Taxonomy.TaxonomyFieldValue' ) {
                       property.$promise = this.resolveTerm(property.TermGuid).then( Function.createDelegate(property,function(term){
                          this.Term = term;
                          this.Label = term.Name;
                       }));
                   }
                }
             }
          }
          this.resolveAll(done);
    }
  };
  } ] );

ngSharePoint.factory("ListResource", ['$http','$resource', 'SharePoint', function ($http, $resource, sharePoint) {
           function transformSharePointRESTRequest(data, headersGetter) {
              if ( typeof data!== 'undefined' ) {
	              if ( typeof data.web !== 'undefined' ) {
	                delete data.web;
	              }
	              if ( typeof data.listName !== 'undefined' ) {
	                delete data.listName;
	              }
	              if ( typeof data.camlQuery !== 'undefined' ) {
	                delete data.camlQuery;
	              }
	              var json = angular.toJson(data);
	              if ( json === "{}" ) {
	                 return "";
	              }
	              else {
	                 return json;
	              }
              }
           }
           function transformSharePointRESTResponse( data, headerGetter ) {
              var raw = null;
              try
              {
                 raw = angular.fromJson(data);
              }
              catch(err) {
                 throw 'list url did not return json (is the list name correct)?';
              }
              if ( typeof raw.d === 'object' ) {
                return raw.d.results;
              }
              else if ( typeof raw.error === 'object' ) {
                 throw raw.error;
              }
              else {
                 return [];
              }
           }
           return $resource(
             sharePoint.getSiteEscaped() + ":web/_api/web/lists/getByTitle(':listName')/items",
           { 
           },
           { 
             query: {
               url: sharePoint.getSiteEscaped() + ":web/_api/web/lists/getByTitle(':listName')/items",
               params: {
                  web: sharePoint.getWeb,
	              listName: "@listName"
               },
               method: 'GET',
               isArray: true,
               headers: { 'Accept': 'application/json; odata=verbose' },
               transformRequest: transformSharePointRESTRequest,
               transformResponse: transformSharePointRESTResponse                   
             },
             camlQuery: {
                method:'POST',
                url: sharePoint.getSiteEscaped() + ':web/_api/web/lists/getByTitle(%27:listName%27)/getItems(query=@v1)?@v1={"ViewXml":"<View><Query>:camlQuery</Query></View>"}',
                params: {
                  web: sharePoint.getWeb,
                  listName: "@listName",
                  camlQuery:  '@camlQuery'
                },
                isArray:true,
                headers: { 'Accept': 'application/json;odata=verbose,application/json;charset=utf-8',
                           'X-RequestDigest': $("#__REQUESTDIGEST").val() 
                         },
                transformRequest: transformSharePointRESTRequest,
                transformResponse: transformSharePointRESTResponse
             },
             update: {
                method:'PUT'
             }
           }    
          );
     }]);
ngSharePoint.directive('spSharepoint', ['SharePoint', function (sharePoint) {
     return {
         restrict: 'AE',
         scope: false,
         controllerAs: 'spSharepoint',
         controller: function($scope) {
            var displayMode = sharePoint.getDisplayMode();
            if ( displayMode !== 'Browse' ) {
               throw "Sharepoint is in " + displayMode + " mode (angular prevented from running).";
            }
         },
         compile: function() {
            return {
              pre: function preLink(scope, elem, attrs){
		               if ( typeof attrs.site === 'string') {
		                  sharePoint.setSite( attrs.site );
		               }
		               if ( typeof attrs.web === 'string') {
		                   sharePoint.setWeb( attrs.web );
		              }
	               }
            };
         }
     };
}]);
ngSharePoint.directive('spSite', ['SharePoint', function (sharePoint) {
     return {
         restrict: 'AE',
         scope: false,
         link: function preLink( scope, elem, attrs) {
             var site = attrs.spSite;
             if ( typeof scope.spSharepoint !== 'undefined' ) {
                scope.spSharepoint.site = site;
             }
             sharePoint.setSite(site);
         }
      };
}]);
ngSharePoint.directive('spWeb', ['SharePoint', function (sharePoint) {
     return {
         restrict: 'AE',
         scope: false,
         link: function preLink( scope, elem, attrs) {
             var web = attrs.spWeb;
             sharePoint.setWeb(web);
             if ( typeof scope.spSharepoint !== 'undefined' ) {
                scope.spSharepoint.web = web;
             }
         }
      };
}]);
ngSharePoint.directive('spDisplayMode', ['SharePoint', function (sharePoint) {
     return {
         restrict: 'AE',
         scope: false,
         controller: function($scope) {
            $scope.displayMode = sharePoint.getDisplayMode();
         },
         link: function(scope, elem, attrs) {
            var mode = attrs.mode;
            if ( typeof mode === "undefined") {
               mode = attrs.spDisplayMode;
            }
            if ( typeof mode === "undefined" || mode === "?") {
               elem.text(scope.displayMode);
            }
            else if ( mode == scope.displayMode ) {
            }
            else {
                elem.empty();
            }
         }
     };
}]);
ngSharePoint.directive('spCamlQuery', ['$parse', function ($parse) {
     return {
         restrict: 'AE',
         scope: false,
         compile: function( elem, attrs ) {
             elem.hide();
             var camlQueryName = attrs.spCamlQuery;
             var template = elem.html();
             if ( template.search(/&lt;/i) >= 0 ) {
                template = template.replace(/&lt;/gi,"<").replace(/&gt;/gi,">");
             }
             
	         var templateProvider = new TextTemplateProvider($parse,'text/caml')
             template = templateProvider.normalizeSpace(template);
	         var camlTemplate = templateProvider.template(template);
             return {
		         post: function preLink( scope, elem, attrs) {
			         var camlQuery = camlTemplate(scope);
			         if ( camlQueryName.length > 0 ) {
				         scope[camlQueryName] = camlQuery;
			         }
		         }
         	};
         }
      };
}]);
function TextTemplateProvider($parse, templateType) {
	function TextBinding(camlText) {
	  var text = camlText;
	  return function(parser, scope) {
	    return text;
	  }
	}
	
	function ExpressionBinding(camlExpression) {
	  var expression = camlExpression;
	  return function(parser, scope) {
		 return parser(expression)(scope);
	  }
	}
	
	function normalizeSpace ( text ) {
		 // Replace repeated spaces, newlines and tabs with a single space
		 return text.replace(/^\s*|\s(?=\s)|\s*$/g, "");
    }
    
	function getInlineTemplate( templateId) {
		var $templateElement = angular.element('#'+templateId);
		if ( $templateElement .length === 0 ) {
			throw 'In-line template id="' + templateId + '" was not found.';
		}
		else if ( $templateElement .attr('type') !== templateType ) {
			throw 'In-line template id="' + templateId + '" is not of type="' + templateType + '".';
		}
		return normalizeSpace($templateElement.html());
	}
	return {
	    normalizeSpace: normalizeSpace,
		template: function( templateOrId ) {
		   if ( typeof templateOrId !== 'string') {
		       throw "Missing template or templateId";
		   }
		   if ( templateOrId.indexOf('<') < 0) {
		       //user supplied the templateId, the template should be delcared in a <script> tag
		       templateId = templateOrId;
		       template = getInlineTemplate( templateOrId );
		   }
		   else {
		      //user supplied the actual template
		      templateId = "anonymous";
		      template = templateOrId;
		   }
	       var bindings = [];
		      var start = 0;
		       while(start < template.length) {
		           var sep = template.indexOf('{{',start);
		           if ( sep >= 0 ) {
		              var end = template.indexOf('}}',sep);
		              if ( end >= 0 ) {
		                 var expression = template.substring(sep+2,end);
		                 bindings.push( new TextBinding(template.substring(start,sep)) );
		                 bindings.push( new ExpressionBinding(expression) );
		                 start = end+2;
		              }
		              else {
		                 bindings.push( new TextBinding(template.substring(start,sep+2)) );
		                 start = sep+2;	                 
		              }
		           }
		           else {
	                 bindings.push( new TextBinding(template.substr(start)) );
		             start = template.length;
		           }
		       }
		       return function( scope, context){
	              var camlScope = scope.$new(false);
	              if ( typeof context !== 'undefined' ) {
		              for( var key in context) {
		                camlScope[key] = context[key]; //copy contextual overrides into new caml scope
		              }
	              }
	              var text = "";
	              for( var i in bindings ) {
	                 text += bindings[i]($parse, camlScope);
	              }
	              return text;
		       };
		    }
	   };
}
     
ngSharePoint.directive('spList', ['$parse','$compile','ListResource','TermStoreService','SharePoint', function ($parse, $compile, listResource, termStoreService, sharePoint) {
/*  translate parameter names from _xyz to $zyx */
	function _underscoreToDollar( data ) {
		for( key in data ) {
			if ( key[0] == '_' ) {
				var value = data[key];
				delete data[key];
				data['$'+key.substr(1)] = value;
			}
		}
		return data;
	}

     return {
         restrict: 'AE',
         scope: false,
         controller: function($scope) {
            var me = this;
		    $scope.listName = "";
		    $scope.listState = "undefined";
		    $scope.listMessage = "";
		    $scope.listDefaultOptions = {
		      initialQuery: true, //when true list will be queried automatically, if false you must call queryList()
		      expandManagedMetaData: true //when true all managed meta data fields will be expanded asynchronously (not required when useCamlQuery = true)
		    };
		    $scope.listDefaultParameters = {
		    };
		    $scope.listOptions = angular.extend({},$scope.listDefaultOptions);
		    $scope.listParameters = angular.extend({},$scope.listDefaultParameters);
		    $scope.items = [];
		    $scope.queryList = function( listName, parameters, options ) {
		       me.queryList( this, listName, parameters, options);
		    }
		    $scope.refreshQuery = function( listName, parameters, options ) {
		       me.refreshQuery ( this, parameters);
		    }
		    this.queryList = function( scope, listName, parameters, options ) {
		      scope.listOptions = angular.extend({}, $scope.listDefaultOptions, options );
		      scope.listParameters = angular.extend({listName: listName}, $scope.listDefaultParameters, parameters);
		      scope.listName = $scope.listParameters.listName;
		      scope.listState = "initialized";
		      if ( scope.listOptions.initialQuery ) {
		          return this.refreshQuery(scope);
		      }
		      };
		    this.refreshQuery = function( scope, parameters ) {
		        scope.listState = "loading";
		        var listParameters = _underscoreToDollar(angular.extend({}, scope.listParameters, parameters));
		        scope.listName = listParameters.listName;
		        var resourceResponse;
		        if ( listParameters.camlQuery ) {
		           var templateProvider = new TextTemplateProvider($parse,'text/caml')
		           var camlTemplate = templateProvider.template(listParameters.camlQuery);
		           listParameters.camlQuery = camlTemplate(scope,listParameters);
		           resourceResponse = listResource.camlQuery(listParameters);
		        }
		        else {
		           resourceResponse = listResource.query(listParameters);
		        }
		        resourceResponse.$promise.then( function(value) {
		             //on success
		             scope.items = value;
		             scope.listState = "loaded";
					 console.log('' + value.length + ' items loaded for list ' + scope.listName + '.');
					 if ( scope.listOptions.expandManagedMetaData ) {
			             termStoreService.processAllTaxonomyFields( value, function(){
						    console.log('all taxonomy fields processed for ' + scope.listName);
			                scope.listState = "resolved";
			                scope.$apply();
			             } );
		             }
		         },function( err ) {
		             //on error
	                 scope.listState = "error";
		             if ( typeof err === 'undefined' ) {
		                scope.listMessage = "An unknown error occurred querying list '" + scope.listName + "'";
		             }
		             else if ( typeof err.message === 'undefined' ) {
		                scope.listMessage = err;
		             }
		             else {
		                scope.listMessage = err.message.value;
		             }
		         });
			     console.log(scope.listMessage);
		         return resourceResponse.$promise;
	         };
         },
         link: function postLink( scope, elem, attrs) {
               var listName = attrs.spList;
               var parameters = {};
               if ( typeof scope.spSharepoint !== 'undefined' ) {
	               if ( typeof scope.spSharepoint.web !== 'undefined' ) {
	                  parameters.web = scope.spSharepoint.web;
	               }
	               if ( typeof scope.spSharepoint.site !== 'undefined' ) {
	                  parameters.site = scope.spSharepoint.site;
	               }
               }
               if ( typeof attrs.parameters === 'string') {
                 var parametersGetter = $parse(attrs.parameters);
                 angular.extend(parameters, parametersGetter(scope));
               }

               var options = {};
               if ( typeof attrs.options === 'string') {
                 var optionsGetter = $parse(attrs.options );
                 options = optionsGetter(scope);
               }
               scope.queryList(listName, parameters, options);
         }
     };
} ] );
ngSharePoint.directive('spCalendar', ['$parse','$http','$q','SharePoint', function ($parse, $http, $q, sharePoint) {
var soapXmlTemplate = "<?xml version='1.0' encoding='utf-8'?>" +
  "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" + 
  "<soap:Body>" + 
   "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" + 
      "<listName>{{calendarName}}</listName>" + 
      "<query>" + 
        "<Query>" + 
          "<Where>" + 
            "<DateRangesOverlap>" + 
              "<FieldRefName='EventDate' />" + 
              "<FieldRefName='EndDate' />" + 
              "<FieldRefName='RecurrenceID' />" + 
              "<ValueType='DateTime'>" + 
                "<Month />" + 
              "</Value>" + 
            "</DateRangesOverlap>" + 
          "</Where>" + 
        "</Query>" + 
      "</query>" + 
      "<queryOptions>" + 
        "<QueryOptions>" + 
          "<ExpandRecurrence>TRUE</ExpandRecurrence>" + 
          "<CalendarDate>" + 
            "<Today />" + 
          "</CalendarDate>" + 
          "<ViewAttributesScope='RecursiveAll' />" + 
        "</QueryOptions>" + 
      "</queryOptions>" + 
      "<viewFields>" + 
        "<ViewFields>" + 
          "<FieldRefName='EventDate' />" + 
          "<FieldRefName='EndDate' />" + 
          "<FieldRefName='fAllDayEvent' />" + 
          "<FieldRefName='fRecurrence' />" + 
          "<FieldRefName='Title' />" + 
        "</ViewFields>" + 
      "</viewFields>" + 
    "</GetListItems>" + 
  "</soap:Body>" + 
"</soap:Envelope>";
     function buildSoapXml( scope, calendarName ) {
		var templateProvider = new TextTemplateProvider($parse,'text/xml')
		var camlTemplate = templateProvider.template(soapXmlTemplate);
		var soapXml = camlTemplate(scope,{calendarName:calendarName});
		return soapXml;
     }
     function loadListViaSoap(scope, calendarName) {
            var url = sharePoint.getSite() + sharePoint.getWeb() + "/_vti_bin/lists.asmx";
            var data = buildSoapXml(scope,calendarName);
            var promise = $http.post( url, data, {
              
              headers: { 'Accept': 'text/xml;charset=utf-8,*/*',
                         'Content-Type': 'text/xml;charset=utf-8',
                         //'Content-Type': 'application/soap+xml',
                         'SOAPAction': 'http://schemas.microsoft.com/sharepoint/soap/GetListItems',
                         'X-RequestDigest': $("#__REQUESTDIGEST").val() 
                       },
              transformResponse: function( data, headersGetter, status ) {
                  var items = [];
                  
                  return items;
               }
            });
            return promise ;
     }
     function getQuery() {
       return "";
       return "<Query>" + 
          "<Where>" + 
            "<DateRangesOverlap>" + 
              "<FieldRefName='EventDate' />" + 
              "<FieldRefName='EndDate' />" + 
              "<FieldRefName='RecurrenceID' />" + 
              "<ValueType='DateTime'>" + 
                "<Today />" + 
              "</Value>" + 
            "</DateRangesOverlap>" + 
          "</Where>" + 
        "</Query>";
     }
     function getViewFields(selectColumns) {
        var viewFields =  "<ViewFields>" + 
          "<FieldRefName='EventDate' />" + 
          "<FieldRefName='EndDate' />" + 
          "<FieldRefName='RecurrenceID' />" + 
          "<FieldRefName='RecurrenceData' />" + 
          "<FieldRefName='fAllDayEvent' />" + 
          "<FieldRefName='fRecurrence' />";
          if ( typeof selectColumns === 'undefined' ) {
              selectColumns = 'Title';
          }
          var columns = selectColumns.split(',');
          for( var i in columns ) {
            var column = columns[i].trim();
            viewFields += "<FieldRefName='" + column + "' />";
          }
        viewFields += "</ViewFields>";
        viewFields += "<QueryOptions>";
        viewFields += "<CalendarDate><Today /></CalendarDate>";
        viewFields += "</QueryOptions>";
        return viewFields;
     }
     function getView(selectColumns) {
       var view = '<View Scope="Recursive" RecurrenceRowset="TRUE">' + getQuery() + getViewFields(selectColumns) + '</View>';
       return view ;
     }
     function loadListViaSPClient(scope, calendarName, selectColumns) {
            var deferred = $q.defer();
            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () { 
               var context = SP.ClientContext.get_current();
               var web = context.get_web();
               var calendar = web.get_lists().getByTitle(calendarName);
               if ( calendar ) {
                   var camlQuery = new SP.CamlQuery();
                   camlQuery.set_viewXml( getView(selectColumns) );
                   var listItems = calendar.getItems(camlQuery);
                   context.load(listItems);
                   context.executeQueryAsync(function(){
                       var items = [];
                       var itemEnumerator = listItems.getEnumerator(); 
                       while( itemEnumerator.moveNext() ) {
                         var listItem = itemEnumerator.get_current();
                         items.push( listItem.get_objectData() );
                       }
                   	   deferred.resolve(items);
                   },
                   function(sender,args){
				       var err = 'Error loading calendar "'+ calendarName + '": ' + args.get_message();
					   console.log(err );
					   deferred.reject(err );
	               });
               }
               else {
                   deferred.reject('Calendar "' + calendarName + '" was not found.');
               }
            });
            return deferred.promise;
     }
     return {
         restrict: 'AE',
         scope: false,
         controller: function($scope) {
            $scope.items = [];
            $scope.listState = "loading";
            $scope.listMessage = "";
         },
         link: function postLink( scope, elem, attrs) {
            var calendarName = attrs.spCalendar;
            var selectColumns = attrs.select;
            var promise = loadListViaSPClient(scope, calendarName, selectColumns );
            //var promise = loadListViaSoap(scope, calendarName);
            promise.then( function(data) {
               scope.listState = "loaded";
               scope.items = data;
            }, function(reason) {
               scope.listState = "error";
               scope.listMessage = reason;
            });
         }

     };

} ] );

ngSharePoint.directive('spTermSet', ['$parse','TermStoreService', function ($parse, termStoreService) {
     return {
         restrict: 'AE',
         scope: false,
         controller: function($scope) {
		    $scope.termSetPath = "";
		    $scope.termSetState = "undefined";
		    $scope.termSetMessage = null;
		    $scope.terms = [];
		    $scope.defaultOptions = {
		      depth: 1
		    };
		    $scope.loadTermSet = function(scope, termStoreName, termSetName) {
		       var deferred = termStoreService.loadTerms( termStoreName, termSetName ).then( function(result) {
		         scope.termStore = result.termStore;
		         scope.termSet = result.termSet;
		         scope.terms = result.terms;
		         scope.termSetState = "loaded";
		       },function(reason) {
		         scope.state = reason;
		       });
		    };
		    $scope.loadTerms = function(scope, path, options) {
		       options = angular.extend({},$scope.defaultOptions , options);
		       termStoreService.loadTerms( path, options ).then( function(result) {
		         if ( typeof result === 'undefined') {
		            scope.terms = [];
		         }
		         else if ( result instanceof Array) {
		            scope.terms = result;
		         }
		         else {
		            scope.terms = [ result ]
		         }
		         scope.termSetState = "loaded";
		       },function(reason) {
		         scope.terms = [];
		         scope.termSetState = "error";
		         scope.termSetMessage = reason;
		       });
   		    };
        },
     link: function postLink( scope, elem, attrs) {
           var path = attrs.spTermSet;
           var options = {};
           if ( typeof attrs.options === 'string') {
             var optionsGetter = $parse(attrs.options);
             options = optionsGetter(scope);
           }
           scope.loadTerms(scope, path, options);
     }

   };
} ]);
