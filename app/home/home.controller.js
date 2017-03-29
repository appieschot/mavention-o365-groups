(function () {
  'use strict';

  angular.module('office365app')
    .controller('homeController', ['$scope', '$window', 'dataService', homeController]);

  /**
   * Controller constructor
   */
  function homeController($scope, $window, dataService) {
    var vm = this;
    vm.loading = false;
    vm.myGroups = [];
    vm.navigate = navigate;
    vm.createGroup = createGroup;
    vm.creatingGroup = false;
    vm.groupTitle = null;

    activate();

    function activate() {
      loadUserData();
    }
    
    function navigate($event, url) {
      // only process the click if it not occured on a hyperlink
      var $target = jQuery($event.target).parent('a');
      if ($target.length === 0) {
        $window.open(url, '_blank');
      }
    }

    function loadUserData() {
      vm.loading = true;
      vm.myGroups.length = 0;

      dataService.getMyGroups().then(function (myGroups) {
        myGroups.forEach(function (group) {
          vm.myGroups.push(group);
          loadGroupDetails(group);
        });
      }, function (err) {
        console.error(err);
      }).finally(function () {
        vm.loading = false;
      });
    }

    function createGroup() {
      vm.creatingGroup = true; 
      
      dataService.createGroup(vm.groupTitle).then(function(){
          vm.creatingGroup = false;
      });
    }

    function loadGroupDetails(group) {
      dataService.getGroupDetails(group.id, group.odataId).then(function (groupDetails) {
        for (var i = 0; i < vm.myGroups.length; i++) {
          var g = vm.myGroups[i];
          if (g.id !== group.id) {
            continue;
          }

          g.description = groupDetails.description;
          g.unseenCount = groupDetails.unseenCount;
          g.picture = groupDetails.picture;
        }
      }, function (err) {
        console.error(err);
      });
    }
  }

})();