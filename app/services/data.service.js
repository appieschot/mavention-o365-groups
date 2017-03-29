(function (){
  'use strict';

  angular.module('office365app')
    .service('dataService', ['$http', '$q', dataService]);

  function dataService($http, $q){
    var api = 'https://graph.microsoft.com/v1.0/';
    var api_beta = 'https://graph.microsoft.com/beta/'    
    return {
      getMyGroups: getMyGroups,
      getGroupDetails: getGroupDetails,
      getMyPlans: getMyPlans,
      getPlanDetails: getPlanDetails,
    };

    /* ************************************************************ */

    function getMyPlans() {
      var deferred = $q.defer();

      $http({
        url: api_beta + 'me/plans',
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata.metadata=full'
        }
      }).success(function (data) {
        var myPlans = [];
         data.value.forEach(function (planInfo) {
            myPlans.push({
              id: planInfo.id,
              odataId: planInfo['@odata.id'],
              owner: planInfo.owner, 

              displayName: planInfo.title,
              //description: planInfo.description
            });
         });

        deferred.resolve(myPlans);
      }).error(function (err) {
        deferred.reject(err);
      });

      return deferred.promise;
    }
  
    function getPlanDetails(groupId, groupODataId) {
      var deferred = $q.defer();

      $q.all({
        // doesn't work in v1.0
        // unseenCount: getGroupUnseenCount(groupODataId),
        picture: getGroupPicture(groupId),
        filesUrl: getGroupFilesUrl(groupId)
      }).then(function (value) {
        deferred.resolve({
          // unseenCount: value.unseenCount,
          picture: value.picture,
          filesUrl: value.filesUrl
        });
      }, function (err) {
        deferred.reject(err);
      })

      return deferred.promise;
    }

    function getMyGroups() {
      var deferred = $q.defer();

      $http({
        url: api + 'me/memberOf?$top=500',
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata.metadata=full'
        }
      }).success(function (data) {
        var myGroups = [];

        data.value.forEach(function (groupInfo) {
          // workaround as the rest filter for unified groups doesn't seem to work client-side
          if (groupInfo.groupTypes &&
              groupInfo.groupTypes.indexOf('Unified') > -1) {
            myGroups.push({
              id: groupInfo.id,
              odataId: groupInfo['@odata.id'],
              displayName: groupInfo.displayName,
              description: groupInfo.description,
              email: groupInfo.mail,
              // not supported in v1.0
              // isFavorite: groupInfo.isFavorite,
              conversationsUrl: 'https://outlook.office365.com/owa/#path=/group/' + groupInfo.mail + '/mail',
              calendarUrl: 'https://outlook.office365.com/owa/#path=/group/' + groupInfo.mail + '/calendar'
            });
          }
        }, this);
       
        deferred.resolve(myGroups);
      }).error(function (err) {
        deferred.reject(err);
      });

      return deferred.promise;
    }

    function getGroupDetails(groupId) {
      var deferred = $q.defer();

      $q.all({
        // doesn't work in v1.0
        // unseenCount: getGroupUnseenCount(groupODataId),
        picture: getGroupPicture(groupId),
        filesUrl: getGroupFilesUrl(groupId)
      }).then(function (value) {
        deferred.resolve({
          // unseenCount: value.unseenCount,
          picture: value.picture,
          filesUrl: value.filesUrl
        });
      }, function (err) {
        deferred.reject(err);
      })

      return deferred.promise;
    }

    function getGroupPicture(groupId) {
      var deferred = $q.defer();

      $http({
        url: api + 'groups/' + groupId + '/photo/$value',
        method: 'GET',
        responseType: 'blob'
      }).success(function (image) {
        var url = window.URL || window.webkitURL;
        deferred.resolve(url.createObjectURL(image));
      }).error(function (err) {
        deferred.reject(err);
      });

      return deferred.promise;
    }
    
    function getGroupFilesUrl(groupId) {
      var deferred = $q.defer();

      $http({
        url: api + 'groups/' + groupId + '/drive/root?$select=webUrl',
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata.metadata=none'
        }
      }).success(function (data) {
        deferred.resolve(data.webUrl);
      }).error(function (err) {
        deferred.reject(err);
      });

      return deferred.promise;
    }

  }

})();