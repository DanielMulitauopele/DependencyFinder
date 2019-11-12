require './dependency_finder'

df = DependencyFinder.new

df.write_to_excel("opportunity_field_names.txt", [1, 2, 3, 4, 5])