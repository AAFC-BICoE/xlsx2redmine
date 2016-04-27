#!/usr/bin/env python

import yaml
import os
import optparse
from redmine import Redmine
import openpyxl
from datetime import datetime
import logging

def main():
  logging.basicConfig(level=logging.INFO, format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s')
  logging.getLogger("requests").setLevel(logging.WARNING)
  logging.getLogger("urllib3").setLevel(logging.WARNING)

  # Convert logging to stream to capture initial debug messages
  logging.info('Loading and configuring options parser.')
  parser = optparse.OptionParser()
  parser.add_option('-c', '--configuration', dest='config_file', help='Configuration file path [Required]')
  parser.add_option('-v', '--verbose', action="store_true", dest='debug', help='Enable verbose output.')
 
  logging.debug('Parsing command line options.')
  (options, args) = parser.parse_args()

  if options.debug:
    logging.getLogger().setLevel(logging.DEBUG)

  if not options.config_file:
    parser.error('Configuration file option (-c, --configuration) is required.')

  config = load_config_file(options.config_file)


  rmc = config['redmine']
  try:
    if 'api-key' in rmc and rmc['api-key']:
      redmine = Redmine(rmc['url'], key=rmc['api-key'], requests={'verify': False})
    elif 'username' in rmc and 'password' in rmc and rmc['username'] and rmc['password']:
      redmine = Redmine(rmc['url'], username=rmc['username'], password=rmc['password'], requests={'verify': False})
    else:
      raise Exception('You must provide the username/password fields or the api-key field in the configuration file.')
  except AuthError:
    logging.error('Redmine authentication failed.  Check your username and password and try again.')
    raise

  logging.info('Completed Redmine authentication.')

  tasks = parse_tasks(redmine,
                      config['project']['spreadsheet']['path'], 
                      config['project']['spreadsheet']['sheet-name'], 
                      config['project']['spreadsheet']['map'], 
                      config['project']['id'], 
                      config['project']['tracker-id'])

  logging.info('Begining import of tasks into redmine.')
  issues_created = 0
  for task in tasks.values():
    task.create_issue()
    task.create_predecation(tasks)
    if task.issue_id:
      issues_created += 1

  logging.info('Created {} issues.'.format(issues_created))
  logging.info('Completed importing tasks.')

  

def load_config_file(path):
  logging.debug('Validating if configuration file path exists {}'.format(path))
  if not os.path.isfile(path):
    logging.error('File not found {}'.format(path))
    raise IOError('File not found: {}'.format(path))

  logging.debug('Loading configuration file {}'.format(path))
  with open(path, 'r') as config_file:
    return yaml.load(config_file.read())

  logging.error('Failed to load configuration file {}.'.format(path))

def parse_tasks(redmine, workbook_path, sheet_name, mapping, project_id, tracker_id):
  logging.info('Parsing excel workbook {} for tasks.'.format(workbook_path))
  if not os.path.isfile(workbook_path):
    logging.error('Excel document does not exist at {}.'.format(workbook_path))
    return None

  wb = openpyxl.load_workbook(workbook_path, data_only=True)
  logging.info('Loaded excel document {}.'.format(workbook_path))
  sheet = wb.get_sheet_by_name(sheet_name)

  logging.info('Parsing tasks from workbook sheet {} starting at row 2.'.format(sheet_name))
  task_list = {}
  for row in range(2, sheet.max_row + 1):
    logging.debug('Parsing row {}'.format(row))

    task = Task(redmine)
    task.id = sheet[mapping['id'] + str(row)].value
    task.project_id = project_id
    task.tracker_id = tracker_id
    task.subject = sheet[mapping['subject'] + str(row)].value
    task.assignee = sheet[mapping['assignee'] + str(row)].value
    task.wbs = sheet[mapping['wbs'] + str(row)].value
    
    predecessors = sheet[mapping['predecessor'] + str(row)].value
    if predecessors is not None:
      task.predecessor_ids = str(predecessors).split(',')

    task.start_date = sheet[mapping['start-date'] + str(row)].value.date()
    task.due_date = sheet[mapping['due-date'] + str(row)].value.date()

    logging.debug('Created task {}'.format(task))

    task_list[task.id] = task

  logging.info('Setting each task\'s parent-child relationships')
  for task in task_list.values():
    task.parent_task = get_parent_task(task_list, task)

  logging.info('Loaded {} tasks from workbook.'.format(len(task_list)))
  
  return task_list

def get_parent_task(task_list, task):
  if task.wbs is None:
    return None

  parent_wbs = '.'.join(task.wbs.split('.')[0:-1])
  if parent_wbs is None or parent_wbs == '':
    logging.debug('Skipping parent task search because task {} is a top level task.'.format(task.id))
    return None

  logging.debug('Searching task list for task {}\'s parent with WBS {}'.format(task.id, parent_wbs))
  for t in task_list.values():
    if t.wbs == parent_wbs:
      return t
  logging.warning('Did not find parent task for task {}.'.format(task.id))
  return None

class Task():
  def __init__(self, redmine):
    self.redmine = redmine
    self.project_id = None
    self.parent_task = None
    self.subject = None
    self.tracker_id = None
    self.description = None
    self.start_date = None
    self.due_date = None
    self.issue_id = None
    self.assignee_id = None
    self.assignee = None
    self.predecessor_ids = []

  def __str__(self):
    return '''Project ID: {}, Parent ID: {}, Subject: {}, Tracker ID: {}, Description: {},
            Start Date: {}, Due Date: {}, Issue ID: {}, Assignee ID: {}, Assignee: {}, 
            Predecessors: {}'''\
            .format(self.project_id, self.parent_task, self.subject, self.tracker_id, self.description,
                    self.start_date, self.due_date, self.issue_id, self.assignee_id, self.assignee,
                    self.predecessor_ids)

  def create_issue(self):
    logging.debug('Creating issue for task {}.'.format(self.id))
    # Do not create an issue that we already have an ID for to avoid duplicates
    if self.issue_id is not None:
      logging.debug('Task {} already has an issue ID ({})'.format(self.id, self.issue_id))
      return self.issue_id

    # Ensure that the assignee id is available before creating the issue
    if self.assignee and self.assignee_id is None:
      self.get_assignee_id()

    issue = self.redmine.issue.new()
    issue.project_id = self.project_id
    issue.subject = self.subject
    issue.tracker_id = self.tracker_id
    issue.description = self.description
    issue.start_date = self.start_date
    issue.due_date = self.due_date
    issue.assigned_to_id = self.assignee_id

    # Set the parent issue id for this issue but ensure the parent issue exists first
    if self.parent_task:
      logging.debug('Setting parent-child relationship for issue of task {} to issue of task {}.'.format(self.parent_task.id, self.id))
      if self.parent_task.issue_id is None:
        logging.debug('Task {} has a parent\'s ({}) with no issue ID.  Creating issue before setting parent-child relationship.'.format(self.id, self.parent_task.id))
        self.parent_task.create_issue()
      issue.parent_issue_id = self.parent_task.issue_id

    saved = issue.save()
  
    if saved:
      self.issue_id = issue.id
      logging.info('Created issue {} for task {}'.format(self.issue_id, self.id))
      return self.issue_id

    logging.warning('Failed to create an issue for task {}'.format(self.id))
    return None

  def create_predecation(self, task_list):
    logging.debug('Adding precedessor relationships for task {}'.format(self.id))

    # If this task does not have an issue ID, then it needs to be created first
    if self.issue_id is None:
      logging.debug('Task {} issue ID is None. Creating issue before assigning predecessors.'.format(self.id))
      self.create_issue(redmine)

    for predecessor_id in self.predecessor_ids:
      predecessor_task = task_list[predecessor_id]
      # Create the issue to get the issue id before trying to create a relationship
      if predecessor_task.issue_id is None:
        logging.debug('Task {}\'s predecessor ({}) does not have an issue ID.'.format(self.id, predecessor_task.id))
	logging.debug('Creating issue for task {}.'.format(predecessor_task.id))
        predecessor_task.create_issue()
      logging.debug('Adding predecessor to issue {} for {}.'.format(self.issue_id, predecessor_task.issue_id))
      self.redmine.issue_relation.create(issue_id=self.issue_id, issue_to_id=predecessor_task.issue_id, relation_type='follows')
      logging.info('Added predecessor to issue {} for {}.'.format(self.issue_id, predecessor_task.issue_id))

  def get_assignee_id(self):
    if self.assignee:
      logging.debug('Task {} has assignee {}. Fetching assignee ID.'.format(self.id, self.assignee))
      user_results = self.redmine.user.filter(name=self.assignee)
      user_results._evaluate()
      
      if user_results.total_count > 0:
        self.assignee_id = user_results[0].id
        logging.debug('Obtained assignee ID {} for {}.'.format(self.assignee_id, self.assignee))
	return self.assignee_id
      
      logging.warning('Unable to obtain ID for assignee {}'.format(self.assignee))

    logging.debug('Task {} has no assignee.'.format(self.id))
    return None

if __name__ == '__main__':
    main()

